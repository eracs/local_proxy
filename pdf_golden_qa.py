#!/usr/bin/env python3
"""Generate golden Q&A pairs from a PDF via a Dify workflow.

Usage:
    python pdf_golden_qa.py input.pdf
    python pdf_golden_qa.py input.pdf -o out.jsonl --chunk-chars 2000 --qa-per-chunk 5
    python pdf_golden_qa.py input.pdf --max-chunks 2   # smoke test

Pipeline:
    1. Extract text per page with pypdf.
    2. Pack paragraphs into character-bounded chunks (each chunk tracks its pages).
    3. For each chunk, run the Dify workflow once: sys_prompt carries the QA
       generation instructions, user_prompt carries the chunk text.
    4. Extract the JSON array of QA pairs from the workflow output and append
       them to a JSONL file incrementally.

Dify workflow contract (docs: https://docs.dify.ai/zh/api-reference/guides/workflow):
    POST {BASE_URL}/workflows/run
    headers:  Authorization: Bearer <API_KEY>
    body:     {"inputs": {"sys_prompt": ..., "user_prompt": ...},
               "response_mode": "blocking", "user": ...}
    response: {"data": {"status": "succeeded", "outputs": {"text": ...}, ...}}

Output JSONL record:
    {"question": ..., "answer": ..., "pdf": ..., "chunk_id": ..., "pages": [start, end]}
"""

import argparse
import json
import re
import sys
import time
from pathlib import Path

import requests
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Dify configuration -- edit these globals when the server or app changes.
# ---------------------------------------------------------------------------
BASE_URL = "https://dify.h.rua.today/v1"
API_KEY = "app-1JelLoJMX2TNcxAWZygzYqkF"
WORKFLOW_USER = "pdf-golden-qa"  # identity Dify uses for rate limiting / logs

REQUEST_TIMEOUT = 180  # seconds per workflow run
MAX_RETRIES = 3
RETRY_BACKOFF = 5  # seconds, multiplied by attempt number

DEFAULT_CHUNK_CHARS = 2000
DEFAULT_QA_PER_CHUNK = 5
MIN_PARAGRAPH_CHARS = 30  # skip page numbers, running headers, etc.

SYS_PROMPT_TEMPLATE = """You are an expert data annotator creating "golden" question-answer pairs \
for evaluating retrieval / RAG systems.

Given a passage extracted from a PDF, generate {n} high-quality QA pairs.

Rules:
- Every question must be answerable from the passage alone; do not use outside knowledge.
- Answers must be accurate, concise and self-contained.
- Cover the key facts of the passage: names, numbers, dates, definitions, methods, cause-effect.
- Do not refer to "the passage" or "the text" in the questions or answers.
- Vary question types (fact lookup, reason, comparison, definition) and avoid near-duplicates.

Output ONLY a JSON array (no markdown fences, no commentary), each element exactly:
{{"question": "...", "answer": "..."}}"""


class DifyFatalError(Exception):
    """Non-retryable API error (bad request, invalid key, ...)."""


# ---------------------------------------------------------------------------
# Dify workflow client
# ---------------------------------------------------------------------------
def run_workflow(sys_prompt: str, user_prompt: str) -> dict:
    """Run the workflow once; return its `data` object (contains outputs/status)."""
    url = f"{BASE_URL}/workflows/run"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "inputs": {"sys_prompt": sys_prompt, "user_prompt": user_prompt},
        "response_mode": "blocking",
        "user": WORKFLOW_USER,
    }
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        if attempt > 1:
            time.sleep(RETRY_BACKOFF * (attempt - 1))
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)
        except requests.RequestException as e:
            last_err = f"network error: {e}"
            continue
        if resp.status_code >= 500:
            last_err = f"HTTP {resp.status_code}: {resp.text[:300]}"
            continue
        if resp.status_code != 200:
            raise DifyFatalError(f"HTTP {resp.status_code}: {resp.text[:300]}")
        try:
            data = resp.json()["data"]
        except (ValueError, KeyError) as e:
            last_err = f"unexpected response shape ({e}): {resp.text[:300]}"
            continue
        if data.get("status") != "succeeded":
            last_err = f"workflow status={data.get('status')}, error={data.get('error')}"
            continue
        return data
    raise RuntimeError(f"workflow failed after {MAX_RETRIES} attempts: {last_err}")


# ---------------------------------------------------------------------------
# Output parsing
# ---------------------------------------------------------------------------
_JSON_DECODER = json.JSONDecoder()


def extract_qa_items(text: str) -> list:
    """Pull QA pairs out of the workflow's text output.

    The LLM behind the workflow is a reasoning model whose output may embed a
    chain of thought (e.g. a `<!--dify-deepseek-reasoning-->` prefix) before
    the final JSON, so we scan for every decodable JSON array and keep the one
    that holds the most valid QA items. Falls back to bare JSON objects.
    """

    def as_qa(obj) -> dict | None:
        if not isinstance(obj, dict):
            return None
        q = str(obj.get("question", "")).strip()
        a = str(obj.get("answer", "")).strip()
        if q and a:
            return {"question": q, "answer": a}
        return None

    arrays = []
    for m in re.finditer(r"\[", text):
        try:
            obj, _ = _JSON_DECODER.raw_decode(text, m.start())
        except json.JSONDecodeError:
            continue
        if isinstance(obj, list):
            arrays.append([item for item in obj if as_qa(item)])

    if arrays:
        best = max(arrays, key=len)
        if best:
            return best

    items = []
    for m in re.finditer(r"\{", text):
        try:
            obj, _ = _JSON_DECODER.raw_decode(text, m.start())
        except json.JSONDecodeError:
            continue
        qa = as_qa(obj)
        if qa:
            items.append(qa)
    return items


# ---------------------------------------------------------------------------
# PDF text extraction and chunking
# ---------------------------------------------------------------------------
def extract_paragraphs(pdf_path: Path) -> list:
    """Return [(page_no, paragraph_text)] with whitespace normalized."""
    reader = PdfReader(str(pdf_path))
    paragraphs = []
    for page_no, page in enumerate(reader.pages, start=1):
        text = (page.extract_text() or "").strip()
        if not text:
            continue
        for para in re.split(r"\n\s*\n", text):
            para = re.sub(r"\s+", " ", para).strip()
            if len(para) >= MIN_PARAGRAPH_CHARS:
                paragraphs.append((page_no, para))
    return paragraphs


def pack_chunks(paragraphs: list, chunk_chars: int) -> list:
    """Greedily pack paragraphs into chunks of at most chunk_chars characters."""
    chunks = []
    cur_paras, cur_pages, cur_len = [], [], 0

    def flush():
        nonlocal cur_paras, cur_pages, cur_len
        if cur_paras:
            chunks.append(
                {
                    "chunk_id": len(chunks),
                    "pages": [min(cur_pages), max(cur_pages)],
                    "text": "\n\n".join(f"[page {p}] {t}" for p, t in cur_paras),
                }
            )
        cur_paras, cur_pages, cur_len = [], [], 0

    for page_no, para in paragraphs:
        while len(para) > chunk_chars:
            flush()
            chunks.append(
                {
                    "chunk_id": len(chunks),
                    "pages": [page_no, page_no],
                    "text": f"[page {page_no}] " + para[:chunk_chars],
                }
            )
            para = para[chunk_chars:]
        if cur_len + len(para) > chunk_chars and cur_paras:
            flush()
        cur_paras.append((page_no, para))
        cur_pages.append(page_no)
        cur_len += len(para) + 2
    flush()
    return chunks


# ---------------------------------------------------------------------------
# QA generation
# ---------------------------------------------------------------------------
def generate_qa_for_chunk(chunk: dict, qa_per_chunk: int) -> tuple:
    sys_prompt = SYS_PROMPT_TEMPLATE.format(n=qa_per_chunk)
    data = run_workflow(sys_prompt, chunk["text"])
    text = data.get("outputs", {}).get("text") or ""
    items = extract_qa_items(text)
    return items, data


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Generate golden Q&A pairs from a PDF via a Dify workflow.",
    )
    ap.add_argument("pdf", help="input PDF file")
    ap.add_argument("-o", "--output", help="output JSONL path (default: <pdf>.qa.jsonl)")
    ap.add_argument("--chunk-chars", type=int, default=DEFAULT_CHUNK_CHARS,
                    help=f"max characters per chunk (default {DEFAULT_CHUNK_CHARS})")
    ap.add_argument("--qa-per-chunk", type=int, default=DEFAULT_QA_PER_CHUNK,
                    help=f"QA pairs requested per chunk (default {DEFAULT_QA_PER_CHUNK})")
    ap.add_argument("--max-chunks", type=int, default=0,
                    help="only process the first N chunks (0 = all; handy for smoke tests)")
    args = ap.parse_args()

    pdf = Path(args.pdf)
    if not pdf.is_file():
        print(f"error: no such file: {pdf}", file=sys.stderr)
        return 2

    paragraphs = extract_paragraphs(pdf)
    if not paragraphs:
        print("error: no extractable text in this PDF", file=sys.stderr)
        return 2
    chunks = pack_chunks(paragraphs, args.chunk_chars)
    if args.max_chunks > 0:
        chunks = chunks[: args.max_chunks]

    out_path = Path(args.output) if args.output else pdf.with_suffix(".qa.jsonl")
    print(f"{pdf.name}: {len(paragraphs)} paragraphs -> {len(chunks)} chunks "
          f"-> {out_path}", flush=True)

    total_qa, failed_chunks, total_tokens = 0, 0, 0
    with out_path.open("w", encoding="utf-8") as out:
        for i, chunk in enumerate(chunks, start=1):
            t0 = time.time()
            try:
                items, data = generate_qa_for_chunk(chunk, args.qa_per_chunk)
            except DifyFatalError as e:
                print(f"fatal: {e}", file=sys.stderr)
                return 1
            except RuntimeError as e:
                failed_chunks += 1
                print(f"[{i}/{len(chunks)}] pages {chunk['pages'][0]}-{chunk['pages'][1]}: "
                      f"FAILED ({e})", flush=True)
                continue
            total_tokens += data.get("total_tokens") or 0
            for it in items:
                rec = {
                    "question": it["question"],
                    "answer": it["answer"],
                    "pdf": pdf.name,
                    "chunk_id": chunk["chunk_id"],
                    "pages": chunk["pages"],
                }
                out.write(json.dumps(rec, ensure_ascii=False) + "\n")
            out.flush()
            total_qa += len(items)
            print(f"[{i}/{len(chunks)}] pages {chunk['pages'][0]}-{chunk['pages'][1]}: "
                  f"{len(items)} QA pairs ({time.time() - t0:.1f}s)", flush=True)
            if not items:
                print(f"  warning: no QA parsed from chunk {chunk['chunk_id']}", file=sys.stderr)

    print(f"done: {total_qa} QA pairs, {failed_chunks}/{len(chunks)} chunks failed, "
          f"{total_tokens} tokens -> {out_path}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
