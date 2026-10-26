#!/usr/bin/env python3
"""Analyze a plain-text file with the Dify workflow LLM.

Usage:
    python3 analyze_file.py report.txt -p "总结这份文件的风险点，用中文回答"
    python3 analyze_file.py notes.md -p "提取所有待办事项" -o result.md

File content is inserted directly into the user prompt, so the input must be
plain text / text-variant (txt, md, csv, log, code, ...).
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

import requests

# ---------------------------------------------------------------------------
# Dify workflow configuration -- loaded from config.json next to this file
# (same file the golden-QA scripts use; see config.example.json).
# ---------------------------------------------------------------------------
CONFIG_PATH = Path(__file__).resolve().parent / "config.json"


def _load_dify_config() -> tuple:
    if not CONFIG_PATH.is_file():
        raise SystemExit(
            f"error: config not found: {CONFIG_PATH}\n"
            "copy config.example.json to config.json and fill in base_url / api_key")
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise SystemExit(f"error: invalid JSON in {CONFIG_PATH}: {e}")
    base_url = str(data.get("base_url", "")).strip().rstrip("/")
    api_key = str(data.get("api_key", "")).strip()
    if not base_url or not api_key:
        raise SystemExit(f"error: {CONFIG_PATH} needs non-empty base_url and api_key")
    return base_url, api_key


BASE_URL, API_KEY = _load_dify_config()
WORKFLOW_USER = "file-analyzer"

REQUEST_TIMEOUT = 300
MAX_API_RETRIES = 3
RETRY_BACKOFF = 5

MAX_INPUT_CHARS = 60000

SYS_PROMPT = (
    "You are a careful file-analysis assistant. Answer strictly from the file "
    "content provided in the user message; if the file does not contain the "
    "requested information, say so explicitly instead of inventing it. Quote "
    "short excerpts from the file when they support a conclusion. Reply in the "
    "language of the user's instruction."
)


class DifyFatalError(RuntimeError):
    pass


# The workflow LLM is a reasoning model: its text output embeds a chain-of-thought
# block (leading <think>...</think> with a dify marker) before the final answer.
_REASONING_BLOCK = re.compile(
    r"^\s*<think>\s*(?:<!--\s*dify-deepseek-reasoning\s*-->)?.*?</think>\s*",
    re.DOTALL,
)


def clean_answer(text: str) -> str:
    """Strip the leading reasoning block so only the final answer remains."""
    return _REASONING_BLOCK.sub("", text, count=1).strip()


def run_workflow(sys_prompt: str, user_prompt: str) -> dict:
    """Run the workflow once; return its `data` object (contains outputs/status)."""
    url = f"{BASE_URL}/workflows/run"
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    payload = {
        "inputs": {"sys_prompt": sys_prompt, "user_prompt": user_prompt},
        "response_mode": "blocking",
        "user": WORKFLOW_USER,
    }
    last_err = None
    for attempt in range(1, MAX_API_RETRIES + 1):
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
    raise RuntimeError(f"workflow failed after {MAX_API_RETRIES} attempts: {last_err}")


def read_text_file(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def build_user_prompt(prompt: str, content: str, max_chars: int) -> str:
    if len(content) > max_chars:
        content = content[:max_chars]
        print(
            f"warning: file truncated to {max_chars} characters "
            f"(use --max-chars to adjust)",
            file=sys.stderr,
        )
    return (
        f"FILE CONTENT:\n"
        f'"""\n{content}\n"""\n\n'
        f"ANALYSIS REQUEST:\n{prompt}"
    )


def main() -> int:
    global REQUEST_TIMEOUT
    ap = argparse.ArgumentParser(
        description="Send a plain-text file plus an analysis prompt to the Dify workflow LLM."
    )
    ap.add_argument("file", help="input text file (txt, md, csv, log, code, ...)")
    ap.add_argument("-p", "--prompt", required=True, help="analysis instruction")
    ap.add_argument("-o", "--output", help="save the answer to this file (default: stdout)")
    ap.add_argument("--max-chars", type=int, default=MAX_INPUT_CHARS,
                    help=f"max file characters sent to the LLM (default {MAX_INPUT_CHARS})")
    ap.add_argument("--timeout", type=int, default=REQUEST_TIMEOUT,
                    help=f"per-call timeout in seconds (default {REQUEST_TIMEOUT})")
    args = ap.parse_args()

    path = Path(args.file)
    if not path.is_file():
        print(f"error: no such file: {path}", file=sys.stderr)
        return 2
    REQUEST_TIMEOUT = args.timeout

    content = read_text_file(path)
    if not content.strip():
        print("error: file is empty", file=sys.stderr)
        return 2

    user_prompt = build_user_prompt(args.prompt, content, args.max_chars)
    try:
        data = run_workflow(SYS_PROMPT, user_prompt)
    except (DifyFatalError, RuntimeError) as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    answer = (data.get("outputs") or {}).get("text") or ""
    if not answer.strip():
        print("error: workflow returned empty text", file=sys.stderr)
        return 1

    cleaned = clean_answer(answer)

    if args.output:
        Path(args.output).write_text(cleaned, encoding="utf-8")
        print(f"saved: {args.output}", file=sys.stderr)
    else:
        print(cleaned)
    print(
        f"[tokens={data.get('total_tokens') or '?'} elapsed={data.get('elapsed_time')}s]",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
