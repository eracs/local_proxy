#!/usr/bin/env python3
"""Part 1 of the golden-QA pipeline: generate golden Q&A from a PDF.

Pipeline (framework: goldqa/ master prompts 1-2 + Stage A of prompt 3):
    1. per chunk, ask the LLM for base questions with per-language
       question/answer, enquiry type, category and complexity.
    2. Stage A "Expected Response Review": score each golden answer on the
       five criteria (PASS >= 95%), FAIL triggers up to --max-benchmark-reruns
       targeted repairs, each followed by a fresh re-review.
    3. write a template-conformant xlsx (Golden Q&A Master / Text Execution
       Result / Report) AND a golden_qa.json handoff consumed by
       evaluate_qa.py (part 2, real-system testing).

Usage:
    python3 generate_qa.py 手册.pdf                    # full run
    python3 generate_qa.py 手册.pdf --languages en --max-chunks 1 --qa-per-chunk 1
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

from qa_common import (
    ALL_METRICS, AP_PENDING_MODE, AP_REVIEW_MODE, APPENDED_HEADERS, BOLD,
    DEFAULT_CHUNK_CHARS, DEFAULT_LANGUAGES, DEFAULT_QA_PER_CHUNK,
    DifyFatalError, GEN_SYS_PROMPT, LANG_SPECS, LANG_SUFFIX,
    MAX_BENCHMARK_RERUNS, METRIC_LABELS, PARSE_RETRIES, PROFILE_NAME,
    REPAIR_SYS_PROMPT, RETRY_HINT, REVIEW_SYS_PROMPT, RISK_METRICS, RunLog,
    TEMPLATE_PATH, TER_HEADERS, REQUEST_TIMEOUT, _header_col, _norm_lang_keys,
    _page_str, _report_section, _scored_entry, _write, append_run_columns,
    call_llm, clear_metric_cells, extract_best, extract_paragraphs,
    metric_cells, metric_scores, normalized_quality, overall_of, pack_chunks,
    prepare_sheets, verdict_of, write_master_row,
)


# ---------------------------------------------------------------------------
# Stage 1: generation
# ---------------------------------------------------------------------------
def _valid_lang_case(v) -> bool:
    return (
        isinstance(v, dict)
        and isinstance(v.get("question"), str)
        and isinstance(v.get("answer"), str)
        and v.get("question", "").strip()
        and v.get("answer", "").strip()
    )


def _question_array_validator(obj) -> int:
    if not isinstance(obj, list):
        return 0
    n = 0
    for it in obj:
        if isinstance(it, dict) and any(
            _valid_lang_case(it.get(k)) for k in ("tc", "sc", "en")
        ):
            n += 1
    return n


def generate_base_questions(chunk: dict, n: int, languages: tuple, log: RunLog) -> list:
    sys_prompt = GEN_SYS_PROMPT.format(
        n=n, languages=", ".join(LANG_SPECS[l] for l in languages)
    )
    items = None
    for attempt in range(1, PARSE_RETRIES + 1):
        text, elapsed, tokens = call_llm(
            sys_prompt, chunk["text"] + (RETRY_HINT if attempt > 1 else "")
        )
        arr = extract_best(text, _question_array_validator)
        if isinstance(arr, list) and arr:
            items = arr
            log.event("generation", chunk_id=chunk["chunk_id"], elapsed_ms=elapsed,
                      tokens=tokens, items=len(items), raw_output=text)
            break
        log.event("generation_parse_retry", chunk_id=chunk["chunk_id"], attempt=attempt,
                  raw_output=text[:2000])
    if items is None:
        raise RuntimeError(f"no parseable QA from chunk {chunk['chunk_id']}")

    questions = []
    for it in items:
        norm = _norm_lang_keys(it)
        langs = {k: norm[k] for k in languages if _valid_lang_case(norm.get(k))}
        if not langs:
            continue
        pages = _parse_pages(it.get("pages"), chunk["pages"])
        questions.append(
            {
                "enquiry_type": str(it.get("enquiry_type", "")).strip() or "General",
                "category": str(it.get("category", "")).strip() or "general",
                "complexity": str(it.get("complexity", "")).strip() or "Moderate",
                "answer_format": str(it.get("answer_format", "")).strip() or "guided",
                "pages": pages,
                "chunk_id": chunk["chunk_id"],
                "chunk_pages": chunk["pages"],
                "chunk_text": chunk["text"],
                "cases": langs,
            }
        )
    return questions


def _parse_pages(raw, chunk_pages: list) -> list:
    """Model-reported page list, validated against the chunk's page range."""
    lo, hi = chunk_pages
    nums = []
    if isinstance(raw, str):
        nums = [int(x) for x in re.findall(r"\d+", raw)]
    elif isinstance(raw, (int, float)):
        nums = [int(raw)]
    nums = [p for p in nums if lo <= p <= hi]
    if not nums:
        return [lo, hi]
    return [min(nums), max(nums)]


# ---------------------------------------------------------------------------
# Stage A: Expected Response Review (+ repairs)
# ---------------------------------------------------------------------------
def _review_user_prompt(item: dict, languages: tuple) -> str:
    parts = ["SOURCE PASSAGE (with [page N] markers):", item["chunk_text"], ""]
    parts.append(
        f"BASE QUESTION {item['base_id']} — QUESTION AND EXPECTED RESPONSE UNDER REVIEW "
        f"(language keys in your JSON must be exactly these lowercase codes: "
        f"{', '.join(languages)}):"
    )
    for lang in languages:
        c = item["cases"][lang]
        parts.append(f"- [{lang}] question: {c['question']}")
        parts.append(f"  expected response: {c['answer']}")
    parts.append("")
    parts.append("Score every listed language version.")
    return "\n".join(parts)


def _review_validator(languages: tuple):
    def validate(obj) -> int:
        if not isinstance(obj, dict):
            return 0
        cases = _norm_lang_keys(obj.get("cases"))
        if not cases:
            return 0
        n = 0
        for lang in languages:
            if lang in cases and _scored_entry(cases[lang]):
                n += 1
        return n

    return validate


def review_question(item: dict, languages: tuple, log: RunLog, version: int) -> dict:
    """One review call for all language cases; returns {lang: {raw, overall, ...}}."""
    user_prompt = _review_user_prompt(item, languages)
    validator = _review_validator(languages)
    result = None
    for attempt in range(1, PARSE_RETRIES + 1):
        text, elapsed, tokens = call_llm(
            REVIEW_SYS_PROMPT, user_prompt + (RETRY_HINT if attempt > 1 else "")
        )
        obj = extract_best(text, validator)
        if obj is not None:
            result = obj["cases"]
            log.event("stage_a_review", base_id=item["base_id"], version=version,
                      elapsed_ms=elapsed, tokens=tokens, raw_output=text)
            break
        log.event("stage_a_review_parse_retry", base_id=item["base_id"], version=version,
                  attempt=attempt, raw_output=text[:2000])
    if result is None:
        raise RuntimeError(f"review output unparseable for {item['base_id']}")

    result = _norm_lang_keys(result)
    evaluated = {}
    for lang in languages:
        entry = result.get(lang)
        if not _scored_entry(entry):
            continue
        raw = metric_scores(entry)
        evaluated[lang] = {
            "raw": raw,
            "overall": overall_of(raw),
            "pass": verdict_of(overall_of(raw)) == "PASS",
            "reasons": {
                m: str(entry.get("reasons", {}).get(m, "")).strip()
                for m in ALL_METRICS
            },
        }
    return evaluated


def _repair_user_prompt(item: dict, languages: tuple, review: dict) -> str:
    parts = ["SOURCE PASSAGE (with [page N] markers):", item["chunk_text"], ""]
    parts.append(
        f"BASE QUESTION {item['base_id']} — question per language "
        f"(language keys in your JSON must be exactly these lowercase codes: "
        f"{', '.join(languages)}):"
    )
    for lang in languages:
        parts.append(f"- [{lang}] {item['cases'][lang]['question']}")
    parts.append("")
    parts.append("CURRENT EXPECTED RESPONSE AND REVIEW DEFECTS (regenerate the FAILED languages only):")
    for lang in languages:
        if lang not in review:
            continue
        c = item["cases"][lang]
        r = review[lang]
        status = "PASS" if r["pass"] else "FAIL"
        parts.append(f"- [{LANG_SUFFIX[lang]}] ({status}, overall {r['overall'] * 100:.1f}%)")
        parts.append(f"  current answer: {c['answer']}")
        for m in ALL_METRICS:
            parts.append(f"  {m} = {r['raw'][m]:.2f}: {r['reasons'][m]}")
    return "\n".join(parts)


def _repair_validator(failed_langs: tuple):
    def validate(obj) -> int:
        if not isinstance(obj, dict):
            return 0
        cases = _norm_lang_keys(obj.get("cases"))
        if not cases:
            return 0
        return sum(
            1
            for lang in failed_langs
            if isinstance(cases.get(lang), dict)
            and isinstance(cases[lang].get("answer"), str)
            and cases[lang]["answer"].strip()
        )

    return validate


def repair_question(item: dict, languages: tuple, review: dict, log: RunLog, version: int) -> dict:
    failed = tuple(lang for lang in languages if lang in review and not review[lang]["pass"])
    if not failed:
        return {}
    user_prompt = _repair_user_prompt(item, languages, review)
    obj = extract_best(call_llm(REPAIR_SYS_PROMPT, user_prompt)[0], _repair_validator(failed))
    if obj is None:
        log.event("repair_parse_failed", base_id=item["base_id"], version=version)
        return {}
    revised = {}
    cases = _norm_lang_keys(obj.get("cases"))
    for lang in failed:
        new_answer = (cases.get(lang) or {}).get("answer", "")
        if isinstance(new_answer, str) and new_answer.strip():
            revised[lang] = new_answer.strip()
    log.event("repair", base_id=item["base_id"], version=version, langs=list(revised),
              change_log=str(obj.get("change_log", "")))
    return revised


def stage_a(item: dict, languages: tuple, max_reruns: int, log: RunLog) -> dict:
    """Review -> up to max_reruns repairs with fresh re-review. Returns the
    final per-language review, the rerun count and the first-attempt verdict;
    every version is preserved in the log."""
    review = review_question(item, languages, log, version=1)
    first_pass = all(
        lang in review and review[lang]["pass"] for lang in languages
    )
    reruns = 0
    while True:
        failed = [lang for lang in languages if lang in review and not review[lang]["pass"]]
        if not failed or reruns >= max_reruns:
            break
        reruns += 1
        revised = repair_question(item, languages, review, log, version=reruns + 1)
        if not revised:
            break
        for lang, new_answer in revised.items():
            item["cases"][lang]["answer"] = new_answer
        review = review_question(item, languages, log, version=reruns + 1)
    return {"review": review, "reruns": reruns, "first_pass": first_pass}


# ---------------------------------------------------------------------------
# Workbook writing (golden rows + Stage A display; three-sheet adapter)
# ---------------------------------------------------------------------------
def write_workbook(out_path: Path, template_path: Path, run_id: str, cases: list,
                   languages: tuple, pdf_name: str, run_started: str):
    wb = openpyxl.load_workbook(str(template_path))
    master, ter, report = prepare_sheets(wb)

    m_cols = {h: _header_col(master, h) for h in
              ("Test case ID", "Cateogry", "Question", "Remark",
               "Expected response (AI Summarized)", "Expected reference", "Expected page")}
    for i, case in enumerate(cases):
        write_master_row(master, 4 + i, m_cols, case, pdf_name)

    t_cols = {h: _header_col(ter, h) for h in TER_HEADERS}
    app_cols = append_run_columns(ter, list(APPENDED_HEADERS))

    for i, case in enumerate(cases):
        r = 3 + i
        _write(ter, r, t_cols["Test case ID"], case["test_case_id"])
        _write(ter, r, t_cols["Test Set NO"], case["chunk_id"] + 1)
        _write(ter, r, t_cols["Conversation Round NO"], 1)
        _write(ter, r, t_cols["Enquiry Type"], case["enquiry_type"])
        _write(ter, r, t_cols["Cateogry"], case["category"])
        _write(ter, r, t_cols["Complexity"], case["complexity"])
        _write(ter, r, t_cols["Question"], case["question"])
        _write(ter, r, t_cols["Remark (Internal Use Only)"],
               f"Base {case['base_id']}; lang {case['lang']}; format {case['answer_format']}; "
               f"benchmark reruns {case['reruns']}")
        _write(ter, r, t_cols["Expected response (AI Summarized)"], case["expected_response"])
        _write(ter, r, t_cols["Expected reference"], pdf_name)
        _write(ter, r, t_cols["Expected page"], _page_str(case["pages"]))
        _write(ter, r, t_cols["Source document name"], pdf_name)
        _write(ter, r, t_cols["Formatted golden document name"],
               f"{pdf_name}: {_page_str(case['pages'])}")
        _write(ter, r, t_cols["Latency Time (ms)"], "N/A")

        # expected-only review: the five metric fields display this benchmark's
        # Stage A scores; they are cleared once real-system testing runs
        if case["stage_a_raw"]:
            metric_cells(ter, r, t_cols, case["stage_a_raw"], case["stage_a_reasons"])
        else:
            clear_metric_cells(ter, r, t_cols, "Pending: Stage A review incomplete")
        _write(ter, r, t_cols["Overall Status"], case["stage_a_status"])
        _write(ter, r, t_cols["Fail Reason"],
               "" if case["stage_a_status"] == "PASS" else
               f"Expected Response Review: {case['stage_a_status']}"
               + (f", overall {case['stage_a_overall'] * 100:.1f}%"
                  if case["stage_a_overall"] is not None else ""))
        _write(ter, r, app_cols["Expected Response Overall"], case["stage_a_overall"], "0.00%")
        _write(ter, r, app_cols["Review Type"], AP_REVIEW_MODE)
        _write(ter, r, app_cols["Expected Response Rerun Count"], case["reruns"])
        _write(ter, r, app_cols["Run ID"], run_id)

    # --- Report ---
    report["C1"] = "Mean Normalized Quality"
    report["D1"] = "Assessment Coverage"
    report["E1"] = "Total Request"
    report["F1"] = "Note"
    for cell in report[1]:
        cell.font = BOLD

    total = len(cases)
    for i, m in enumerate(("correctness", "answer_relevancy", "role_adherence",
                           "bias", "toxicity")):
        r = 2 + i
        vals = [normalized_quality(m, c["stage_a_raw"][m]) for c in cases
                if c["stage_a_raw"]]
        _write(report, r, 1, i + 1)
        _write(report, r, 2, METRIC_LABELS[m])
        _write(report, r, 3, sum(vals) / len(vals) if vals else None, "0.00")
        _write(report, r, 4, len(vals) / total if total else None, "0%")
        _write(report, r, 5, total)
        _write(report, r, 6, "risk-based metric; normalized quality = 1 - risk"
               if m in RISK_METRICS else "")

    row = _report_section(report, 8, "Run Metadata", [
        ("Run ID", run_id),
        ("Source PDF", pdf_name),
        ("Run started", run_started),
        ("Languages", ", ".join(LANG_SUFFIX[l] for l in languages)),
        ("Scoring profile", PROFILE_NAME),
    ])
    stage_a_pass = sum(1 for c in cases if c["stage_a_status"] == "PASS")
    assessed = [c["stage_a_overall"] for c in cases if c["stage_a_overall"] is not None]
    row = _report_section(report, row, "Expected Response Review (Stage A)", [
        ("Review type", AP_REVIEW_MODE),
        ("Cases", total),
        ("PASS", stage_a_pass),
        ("FAIL", sum(1 for c in cases if c["stage_a_status"] == "FAIL")),
        ("Evaluation Incomplete",
         sum(1 for c in cases if c["stage_a_status"] == "EVALUATION INCOMPLETE")),
        ("First-attempt PASS", sum(1 for c in cases if c.get("first_attempt_pass"))),
        ("Benchmark reruns used", sum(c["reruns"] for c in cases)),
        ("Mean Expected Response Overall",
         sum(assessed) / len(assessed) if assessed else None, "0.00%"),
        ("Full-batch pass rate (PASS / all)",
         stage_a_pass / total if total else None, "0.00%"),
    ])
    _report_section(report, row, "Three-Response Q&A Evaluation (real system Q&A)", [
        ("Mode", AP_PENDING_MODE),
        ("Readiness", "EVALUATION INCOMPLETE (testing not executed here; run "
                      "evaluate_qa.py with real system Q&A — benchmark PASS is "
                      "not test PASS)"),
    ])
    wb.save(str(out_path))


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def main() -> int:
    global REQUEST_TIMEOUT
    ap = argparse.ArgumentParser(
        description="Generate golden Q&A from a PDF via a Dify workflow and run "
                    "the Expected Response Review (Stage A).",
    )
    ap.add_argument("pdf", help="input PDF file")
    ap.add_argument("-o", "--output", help="output xlsx path (default: <pdf>.golden_qa.xlsx)")
    ap.add_argument("--json", dest="json_out",
                    help="golden QA JSON handoff (default: <pdf>.golden_qa.json)")
    ap.add_argument("--log", help="execution log JSONL path (default: <pdf>.qa_log.jsonl)")
    ap.add_argument("--template", default=str(TEMPLATE_PATH),
                    help=f"template xlsx path (default: {TEMPLATE_PATH})")
    ap.add_argument("--chunk-chars", type=int, default=DEFAULT_CHUNK_CHARS,
                    help=f"max characters per chunk (default {DEFAULT_CHUNK_CHARS})")
    ap.add_argument("--qa-per-chunk", type=int, default=DEFAULT_QA_PER_CHUNK,
                    help=f"base questions per chunk (default {DEFAULT_QA_PER_CHUNK})")
    ap.add_argument("--max-chunks", type=int, default=0,
                    help="only process the first N chunks (0 = all)")
    ap.add_argument("--languages", default=",".join(DEFAULT_LANGUAGES),
                    help="comma-separated case languages: tc, sc, en "
                         f"(default {','.join(DEFAULT_LANGUAGES)})")
    ap.add_argument("--max-benchmark-reruns", type=int, default=MAX_BENCHMARK_RERUNS,
                    help=f"max Expected Response repairs per question "
                         f"(default {MAX_BENCHMARK_RERUNS})")
    ap.add_argument("--timeout", type=int, default=REQUEST_TIMEOUT,
                    help=f"per-call timeout in seconds (default {REQUEST_TIMEOUT})")
    args = ap.parse_args()

    languages = tuple(l.strip().lower() for l in args.languages.split(",") if l.strip())
    invalid = [l for l in languages if l not in LANG_SPECS]
    if invalid:
        print(f"error: unknown languages {invalid}; valid: {list(LANG_SPECS)}", file=sys.stderr)
        return 2
    REQUEST_TIMEOUT = args.timeout

    pdf = Path(args.pdf)
    if not pdf.is_file():
        print(f"error: no such file: {pdf}", file=sys.stderr)
        return 2
    template = Path(args.template)
    if not template.is_file():
        print(f"error: template not found: {template}", file=sys.stderr)
        return 2

    paragraphs = extract_paragraphs(pdf)
    if not paragraphs:
        print("error: no extractable text in this PDF", file=sys.stderr)
        return 2
    chunks = pack_chunks(paragraphs, args.chunk_chars)
    if args.max_chunks > 0:
        chunks = chunks[: args.max_chunks]

    out_xlsx = Path(args.output) if args.output else pdf.with_suffix(".golden_qa.xlsx")
    out_json = Path(args.json_out) if args.json_out else pdf.with_suffix(".golden_qa.json")
    log_path = Path(args.log) if args.log else pdf.with_suffix(".qa_log.jsonl")
    run_id = f"RUN-{datetime.now():%Y%m%d-%H%M}"
    run_started = datetime.now().isoformat(timespec="seconds")
    log = RunLog(log_path, run_id)
    log.event("run_started", pdf=pdf.name, chunks=len(chunks), languages=list(languages),
              profile=PROFILE_NAME)

    print(f"{pdf.name}: {len(paragraphs)} paragraphs -> {len(chunks)} chunks | "
          f"languages {languages} | run {run_id}", flush=True)

    base_questions = []
    for i, chunk in enumerate(chunks, start=1):
        try:
            items = generate_base_questions(chunk, args.qa_per_chunk, languages, log)
        except DifyFatalError as e:
            print(f"fatal: {e}", file=sys.stderr)
            return 1
        except RuntimeError as e:
            print(f"[chunk {i}/{len(chunks)}] FAILED: {e}", flush=True)
            continue
        base_questions.extend(items)
        print(f"[gen {i}/{len(chunks)}] {len(items)} base questions "
              f"(pages {_page_str(chunk['pages'])})", flush=True)

    if not base_questions:
        print("error: no questions generated", file=sys.stderr)
        return 1
    for n, item in enumerate(base_questions, start=1):
        item["base_id"] = f"Q{n:03d}"

    cases = []
    for item in base_questions:
        try:
            outcome = stage_a(item, languages, args.max_benchmark_reruns, log)
        except DifyFatalError as e:
            print(f"fatal: {e}", file=sys.stderr)
            return 1
        except RuntimeError as e:
            print(f"[stage A {item['base_id']}] FAILED: {e}", flush=True)
            outcome = {"review": {}, "reruns": 0, "first_pass": False}

        first_pass = outcome.get("first_pass", False)
        for lang in languages:
            case = {
                "test_case_id": f"{item['base_id']}-{LANG_SUFFIX[lang]}",
                "base_id": item["base_id"],
                "lang": lang,
                "question": item["cases"][lang]["question"],
                "expected_response": item["cases"][lang]["answer"],
                "enquiry_type": item["enquiry_type"],
                "category": item["category"],
                "complexity": item["complexity"],
                "answer_format": item["answer_format"],
                "pages": item["pages"],
                "chunk_id": item["chunk_id"],
                "chunk_pages": item["chunk_pages"],
                "chunk_text": item["chunk_text"],
                "reruns": outcome["reruns"],
                "first_attempt_pass": first_pass,
                "stage_a_raw": None,
                "stage_a_overall": None,
                "stage_a_status": "EVALUATION INCOMPLETE",
                "stage_a_reasons": None,
            }
            review = outcome["review"].get(lang)
            if review:
                case["stage_a_raw"] = review["raw"]
                case["stage_a_overall"] = review["overall"]
                case["stage_a_status"] = verdict_of(review["overall"])
                case["stage_a_reasons"] = review["reasons"]
            cases.append(case)
        summary = []
        for lang in languages:
            review = outcome["review"].get(lang)
            if review:
                summary.append(
                    f"{LANG_SUFFIX[lang]}={verdict_of(review['overall'])}({review['overall']:.1%})"
                )
            else:
                summary.append(f"{LANG_SUFFIX[lang]}=INCOMPLETE")
        print(f"[stage A {item['base_id']}] reruns={outcome['reruns']} "
              + " ".join(summary), flush=True)

    write_workbook(out_xlsx, template, run_id, cases, languages,
                   pdf_name=pdf.name, run_started=run_started)
    out_json.write_text(
        json.dumps({"run_id": run_id, "run_started": run_started, "pdf": pdf.name,
                    "languages": list(languages), "cases": cases},
                   ensure_ascii=False, indent=1),
        encoding="utf-8",
    )

    log.event("run_finished", cases=len(cases),
              stage_a_pass=sum(1 for c in cases if c["stage_a_status"] == "PASS"),
              outputs=[str(out_xlsx), str(out_json)])
    log.close()

    print(f"done: {len(base_questions)} base questions, {len(cases)} cases -> {out_xlsx} "
          f"+ {out_json.name} (log: {log_path})", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
