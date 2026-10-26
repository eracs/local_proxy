#!/usr/bin/env python3
"""Part 2 of the golden-QA pipeline: test/evaluate REAL system Q&A.

Inputs:
    --golden  the golden_qa.json produced by generate_qa.py (benchmark + source)
    --actual  real system Q&A records actually produced by the chatbot:
        .jsonl  one record per line:
                {"question": "...", "answer": "...",
                 "test_case_id": "Q001-EN" (optional direct map),
                 "round": 1 (optional, groups repeat answers into L/M/N),
                 "documents": "observed retrieval" (optional),
                 "latency_ms": 812 (optional), "timestamp": "..."}
        .json   a list of such records (or {"records": [...]})
        .csv    columns question,answer[,test_case_id,round,documents,latency_ms]

Matching a record to the golden bank: explicit test_case_id, else normalized
question equality, else best difflib similarity >= 0.6. Unmatched records get
their own rows marked EVALUATION INCOMPLETE (correctness evidence unavailable).

Evaluation: one Dify call per real answer, scoring Correctness / Role
Adherence / Bias / Toxicity / Answer Relevancy against the source passage and
the golden benchmark (profile EQUAL20_OVERALL95..., PASS >= 95%). A row's
Case Overall is the minimum over its rounds; golden cases with no real answer
stay Pending in the denominator.

Usage:
    python3 evaluate_qa.py --golden 手册.golden_qa.json --actual system_qa.jsonl
"""

import argparse
import csv
import difflib
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

from qa_common import (
    ALL_METRICS, AP_PENDING_MODE, AP_SIT_MODE, APPENDED_HEADERS, BOLD,
    DifyFatalError, EVAL_SYS_PROMPT, LANG_SUFFIX, METRIC_LABELS, PARSE_RETRIES,
    PASS_THRESHOLD, PROFILE_NAME, REQUEST_TIMEOUT, RETRY_HINT, RISK_METRICS,
    RunLog, TEMPLATE_PATH, TER_HEADERS, _header_col, _page_str, _readiness,
    _report_section, _scored_entry, _write, append_run_columns, call_llm,
    clear_metric_cells, extract_best, metric_cells, metric_scores,
    normalized_quality, overall_of, prepare_sheets, verdict_of, write_master_row,
)


# ---------------------------------------------------------------------------
# Inputs
# ---------------------------------------------------------------------------
def load_golden(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    cases = data.get("cases")
    if not isinstance(cases, list) or not cases:
        raise ValueError("golden file has no cases[] (expected generate_qa.py output)")
    return data


def load_actual(path: Path) -> list:
    text = path.read_text(encoding="utf-8-sig")
    records = []
    if path.suffix.lower() == ".jsonl":
        for i, line in enumerate(text.splitlines(), start=1):
            line = line.strip()
            if not line:
                continue
            rec = json.loads(line)
            rec["_order"] = i
            records.append(rec)
    elif path.suffix.lower() == ".json":
        data = json.loads(text)
        if isinstance(data, dict):
            data = data.get("records", [])
        for i, rec in enumerate(data, start=1):
            rec["_order"] = i
            records.append(rec)
    elif path.suffix.lower() == ".csv":
        for i, row in enumerate(csv.DictReader(text.splitlines()), start=1):
            row = {k: v for k, v in row.items() if k and v not in (None, "")}
            row["_order"] = i
            records.append(row)
    else:
        raise ValueError(f"unsupported actual-QA format {path.suffix} (use .jsonl/.json/.csv)")
    clean = []
    for rec in records:
        q = str(rec.get("question", "")).strip()
        a = str(rec.get("answer", "")).strip()
        if not q or not a:
            continue
        rnd = rec.get("round")
        clean.append({
            "question": q,
            "answer": a,
            "test_case_id": str(rec.get("test_case_id", "")).strip() or None,
            "round": int(rnd) if str(rnd).strip().isdigit() and 1 <= int(rnd) <= 3 else None,
            "documents": str(rec.get("documents", "")).strip() or None,
            "latency_ms": rec.get("latency_ms"),
            "timestamp": rec.get("timestamp"),
            "_order": rec["_order"],
        })
    return clean


def _norm_q(q: str) -> str:
    return re.sub(r"\W+", "", q.lower())


def match_golden(rec: dict, golden: dict) -> tuple:
    """Return (golden_case, method). method: id|exact|fuzzy|None."""
    by_id = {c["test_case_id"].lower(): c for c in golden["cases"]}
    if rec["test_case_id"]:
        case = by_id.get(rec["test_case_id"].lower())
        if case:
            return case, "id"
    nq = _norm_q(rec["question"])
    for c in golden["cases"]:
        if _norm_q(c["question"]) == nq:
            return c, "exact"
    best, best_ratio = None, 0.0
    for c in golden["cases"]:
        ratio = difflib.SequenceMatcher(
            None, nq, _norm_q(c["question"])
        ).ratio()
        if ratio > best_ratio:
            best, best_ratio = c, ratio
    if best is not None and best_ratio >= 0.6:
        return best, "fuzzy"
    return None, None


# ---------------------------------------------------------------------------
# Evaluation of one real answer
# ---------------------------------------------------------------------------
def _eval_user_prompt(case: dict, rec: dict) -> str:
    lang = case.get("lang", "")
    return (
        "SOURCE PASSAGE (with [page N] markers):\n"
        f"{case.get('chunk_text') or 'NOT AVAILABLE'}\n\n"
        f"QUESTION [{LANG_SUFFIX.get(lang, lang)}] (golden benchmark question):\n"
        f"{case['question']}\n\n"
        "EXPECTED RESPONSE (benchmark reference — score the REAL SYSTEM ANSWER, "
        f"not the benchmark):\n{case['expected_response']}\n\n"
        f"REAL USER QUESTION ASKED:\n{rec['question']}\n\n"
        f"REAL SYSTEM ANSWER TO EVALUATE:\n{rec['answer']}"
    )


def evaluate_record(case: dict, rec: dict, log: RunLog, row_id: str) -> dict:
    """One evaluation call; returns {raw, overall, pass, reasons} or None."""
    result = None
    for attempt in range(1, PARSE_RETRIES + 1):
        text, elapsed, tokens = call_llm(
            EVAL_SYS_PROMPT,
            _eval_user_prompt(case, rec) + (RETRY_HINT if attempt > 1 else ""),
        )
        obj = extract_best(text, _scored_entry)
        if obj is not None:
            result = obj
            log.event("test_evaluation", row_id=row_id, elapsed_ms=elapsed,
                      tokens=tokens, question=rec["question"],
                      answer=rec["answer"], raw_output=text)
            break
        log.event("test_evaluation_parse_retry", row_id=row_id, attempt=attempt,
                  raw_output=text[:2000])
    if result is None:
        return None
    raw = metric_scores(result)
    overall = overall_of(raw)
    return {
        "raw": raw,
        "overall": overall,
        "pass": verdict_of(overall) == "PASS",
        "reasons": {m: str(result.get("reasons", {}).get(m, "")).strip()
                    for m in ALL_METRICS},
        "rec": rec,
    }


def _fail_reason(row: dict) -> str:
    parts = []
    for k in (1, 2, 3):
        rd = row["rounds"].get(k)
        if k in row["required_rounds"] and rd is None:
            parts.append(f"R{k}: missing/incomplete round")
        elif rd is not None and not rd["pass"]:
            worst = ", ".join(
                f"{m}={normalized_quality(m, v):.2f}"
                for m, v in rd["raw"].items()
                if normalized_quality(m, v) < PASS_THRESHOLD
            )
            parts.append(f"R{k} overall {rd['overall']:.2%} ({worst})")
    if row.get("case_overall") is not None:
        parts.append(f"Case Overall (min) {row['case_overall']:.2%}")
    return "; ".join(parts) or "unknown"


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------
def write_workbook(out_path: Path, template_path: Path, run_id: str, golden: dict,
                   rows: list, pdf_name: str, run_started: str, review_type: str):
    wb = openpyxl.load_workbook(str(template_path))
    master, ter, report = prepare_sheets(wb)

    m_cols = {h: _header_col(master, h) for h in
              ("Test case ID", "Cateogry", "Question", "Remark",
               "Expected response (AI Summarized)", "Expected reference", "Expected page")}
    for i, case in enumerate(golden["cases"]):
        write_master_row(master, 4 + i, m_cols, case, pdf_name)

    t_cols = {h: _header_col(ter, h) for h in TER_HEADERS}
    app_cols = append_run_columns(ter, list(APPENDED_HEADERS))

    for i, row in enumerate(rows):
        r = 3 + i
        case = row.get("case")
        _write(ter, r, t_cols["Test case ID"], row["row_id"])
        _write(ter, r, t_cols["Test Set NO"], (case["chunk_id"] + 1) if case else i + 1)
        _write(ter, r, t_cols["Conversation Round NO"], 1)
        _write(ter, r, t_cols["Enquiry Type"], case["enquiry_type"] if case else "Unknown")
        _write(ter, r, t_cols["Cateogry"], case["category"] if case else "unknown")
        _write(ter, r, t_cols["Complexity"], case["complexity"] if case else "Unknown")
        _write(ter, r, t_cols["Question"],
               case["question"] if case else row["records"][0]["question"])
        _write(ter, r, t_cols["Remark (Internal Use Only)"],
               f"lang {case['lang'] if case else '?'}; real answers {len(row['records'])}"
               + (f"; match {row['match']}" if row.get("match") else "; unmatched"))
        _write(ter, r, t_cols["Expected response (AI Summarized)"],
               case["expected_response"] if case else "NOT AVAILABLE (no golden match)")
        _write(ter, r, t_cols["Expected reference"], pdf_name)
        _write(ter, r, t_cols["Expected page"],
               _page_str(case["pages"]) if case else "N/A")

        slot_hdr = {1: "Actual Answer (Round 1)",
                    2: "Repeat Call Actual Answer (Round 2)",
                    3: "Repeat Call Actual Answer (Round 3)"}
        for slot in (1, 2, 3):
            rd = row["rounds"].get(slot)
            _write(ter, r, t_cols[slot_hdr[slot]], rd["rec"]["answer"] if rd else None)
        first_rec = row["records"][0] if row["records"] else {}
        docs = first_rec.get("documents")
        _write(ter, r, t_cols["Actual Documents"],
               docs if docs else "NOT PROVIDED (real system retrieval not observed)")
        _write(ter, r, t_cols["Source document name"],
               pdf_name if case else "NOT PROVIDED")
        _write(ter, r, t_cols["Formatted golden document name"],
               f"{pdf_name}: {_page_str(case['pages'])}" if case else "NOT PROVIDED")
        lat = first_rec.get("latency_ms")
        _write(ter, r, t_cols["Latency Time (ms)"],
               int(lat) if str(lat).strip().isdigit() else "N/A")

        complete = (row["required_rounds"]
                    and all(row["rounds"].get(k) for k in row["required_rounds"]))
        if complete:
            worst_slot = min(row["required_rounds"],
                             key=lambda k: (row["rounds"][k]["overall"], k))
            worst = row["rounds"][worst_slot]
            metric_cells(ter, r, t_cols, worst["raw"], worst["reasons"])
            _write(ter, r, t_cols["Overall Status"],
                   "PASS" if row["case_overall"] >= PASS_THRESHOLD else "FAIL")
            _write(ter, r, t_cols["Fail Reason"], "" if row["case_overall"] >= PASS_THRESHOLD
                   else _fail_reason(row))
            for slot in (1, 2, 3):
                rd = row["rounds"].get(slot)
                _write(ter, r, app_cols[f"Round {slot} Overall"],
                       rd["overall"] if rd else None, "0.00%")
            _write(ter, r, app_cols["Case Overall"], row["case_overall"], "0.00%")
            _write(ter, r, app_cols["Scored Round"], worst_slot)
            _write(ter, r, app_cols["Review Type"], review_type)
        else:
            if not row["records"]:
                reason = "Pending: no real system answer supplied"
            elif not case:
                reason = ("Incomplete: no golden match — correctness evidence "
                          "unavailable, not evaluated")
            elif not row["rounds"]:
                reason = "Incomplete: not evaluated (round slots full or evaluation failed)"
            else:
                reason = "Incomplete: missing/unevaluated round; " + _fail_reason(row)
            clear_metric_cells(ter, r, t_cols, reason)
            known_fail = any(rd is not None and not rd["pass"]
                             for rd in row["rounds"].values())
            _write(ter, r, t_cols["Overall Status"],
                   "FAIL" if known_fail else
                   ("EVALUATION INCOMPLETE" if row["records"] else "Pending"))
            _write(ter, r, t_cols["Fail Reason"],
                   _fail_reason(row) if known_fail else reason)
            for slot in (1, 2, 3):
                rd = row["rounds"].get(slot)
                _write(ter, r, app_cols[f"Round {slot} Overall"],
                       rd["overall"] if rd else None, "0.00%")
            _write(ter, r, app_cols["Case Overall"], None, "0.00%")
            _write(ter, r, app_cols["Scored Round"], None)
            _write(ter, r, app_cols["Review Type"],
                   review_type if row["records"] else AP_PENDING_MODE)
        _write(ter, r, app_cols["Expected Response Overall"],
               case.get("stage_a_overall") if case else None, "0.00%")
        _write(ter, r, app_cols["Expected Response Rerun Count"],
               case.get("reruns") if case else None)
        _write(ter, r, app_cols["Run ID"], run_id)

    # --- Report ---
    report["C1"] = "Mean Normalized Quality"
    report["D1"] = "Assessment Coverage"
    report["E1"] = "Total Request"
    report["F1"] = "Note"
    for cell in report[1]:
        cell.font = BOLD

    def displayed(row, m):
        req = row["required_rounds"]
        if not req or not all(row["rounds"].get(k) for k in req):
            return None
        worst = min(req, key=lambda k: (row["rounds"][k]["overall"], k))
        return normalized_quality(m, row["rounds"][worst]["raw"][m])

    total = len(rows)
    for i, m in enumerate(("correctness", "answer_relevancy", "role_adherence",
                           "bias", "toxicity")):
        r = 2 + i
        vals = [v for v in (displayed(row, m) for row in rows) if v is not None]
        _write(report, r, 1, i + 1)
        _write(report, r, 2, METRIC_LABELS[m])
        _write(report, r, 3, sum(vals) / len(vals) if vals else None, "0.00")
        _write(report, r, 4, len(vals) / total if total else None, "0%")
        _write(report, r, 5, total)
        note = "risk-based metric; normalized quality = 1 - risk" if m in RISK_METRICS else ""
        _write(report, r, 6, note)

    evaluated_rows = [row for row in rows if row["required_rounds"]
                      and all(row["rounds"].get(k) for k in row["required_rounds"])]
    n_pass = sum(1 for row in evaluated_rows if row["case_overall"] >= PASS_THRESHOLD)
    n_fail = len(evaluated_rows) - n_pass
    n_incomplete = total - len(evaluated_rows)
    unmatched = sum(1 for row in rows if not row.get("case"))
    row = _report_section(report, 8, "Run Metadata", [
        ("Run ID", run_id),
        ("Golden bank", pdf_name),
        ("Run started", run_started),
        ("Scoring profile", PROFILE_NAME),
        ("Review type", review_type),
        ("Golden cases", len(golden["cases"])),
        ("Real records loaded", sum(len(r["records"]) for r in rows)),
        ("Unmatched records", unmatched),
    ])
    row = _report_section(report, row, "Real System Q&A Evaluation", [
        ("Rows (test units)", total),
        ("PASS", n_pass),
        ("FAIL", n_fail),
        ("Evaluation Incomplete / Pending", n_incomplete),
        ("Completion coverage", len(evaluated_rows) / total if total else None, "0%"),
        ("Full-batch pass rate (PASS / all)", n_pass / total if total else None, "0.00%"),
        ("Readiness", _readiness(n_pass, total, n_incomplete)),
    ])
    golden_pass = sum(1 for c in golden["cases"]
                      if c.get("stage_a_status") == "PASS")
    _report_section(report, row, "Golden Benchmark Review (from generate_qa.py)", [
        ("Cases", len(golden["cases"])),
        ("Stage A PASS", golden_pass),
        ("Stage A pass rate", golden_pass / len(golden["cases"])
         if golden["cases"] else None, "0.00%"),
    ])
    wb.save(str(out_path))


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def main() -> int:
    global REQUEST_TIMEOUT
    ap = argparse.ArgumentParser(
        description="Evaluate real system Q&A records against the golden bank "
                    "on the five quality metrics.",
    )
    ap.add_argument("--golden", required=True,
                    help="golden_qa.json produced by generate_qa.py")
    ap.add_argument("--actual", required=True,
                    help="real system Q&A records (.jsonl / .json / .csv)")
    ap.add_argument("-o", "--output",
                    help="output xlsx path (default: <actual>.test_result.xlsx)")
    ap.add_argument("--log", help="execution log JSONL path (default: <actual>.test_log.jsonl)")
    ap.add_argument("--template", default=str(TEMPLATE_PATH),
                    help=f"template xlsx path (default: {TEMPLATE_PATH})")
    ap.add_argument("--review-type", default=AP_SIT_MODE,
                    help=f"review-type label written to AP (default: {AP_SIT_MODE})")
    ap.add_argument("--timeout", type=int, default=REQUEST_TIMEOUT,
                    help=f"per-call timeout in seconds (default {REQUEST_TIMEOUT})")
    args = ap.parse_args()
    REQUEST_TIMEOUT = args.timeout

    golden_path = Path(args.golden)
    actual_path = Path(args.actual)
    for p, what in ((golden_path, "golden"), (actual_path, "actual")):
        if not p.is_file():
            print(f"error: no such {what} file: {p}", file=sys.stderr)
            return 2
    try:
        golden = load_golden(golden_path)
    except (ValueError, json.JSONDecodeError) as e:
        print(f"error: cannot load golden bank: {e}", file=sys.stderr)
        return 2
    try:
        records = load_actual(actual_path)
    except (ValueError, json.JSONDecodeError, csv.Error) as e:
        print(f"error: cannot load actual Q&A: {e}", file=sys.stderr)
        return 2
    if not records:
        print("error: no usable question/answer records in the actual file",
              file=sys.stderr)
        return 2

    template = Path(args.template)
    if not template.is_file():
        print(f"error: template not found: {template}", file=sys.stderr)
        return 2

    out_xlsx = (Path(args.output) if args.output
                else actual_path.with_suffix(".test_result.xlsx"))
    log_path = (Path(args.log) if args.log
                else actual_path.with_suffix(".test_log.jsonl"))
    run_id = f"TEST-{datetime.now():%Y%m%d-%H%M}"
    run_started = datetime.now().isoformat(timespec="seconds")
    log = RunLog(log_path, run_id)
    log.event("run_started", golden=str(golden_path), actual=str(actual_path),
              records=len(records), review_type=args.review_type, profile=PROFILE_NAME)

    # match records -> golden cases; group per case
    groups = {}  # key -> {"case": ..., "match": ..., "records": [...]}
    for rec in records:
        case, method = match_golden(rec, golden)
        key = case["test_case_id"] if case else f"UNMATCHED-{rec['_order']:03d}"
        g = groups.setdefault(key, {"case": case, "match": method, "records": []})
        if case and g["case"] is None:
            g["case"], g["match"] = case, method
        g["records"].append(rec)
    for case in golden["cases"]:  # golden cases without real answers stay pending
        groups.setdefault(case["test_case_id"],
                          {"case": case, "match": None, "records": []})

    rows = []
    golden_order = {c["test_case_id"]: i for i, c in enumerate(golden["cases"])}
    for gi, key in enumerate(sorted(
            groups, key=lambda k: (golden_order.get(k, len(golden_order)), k)), start=1):
        g = groups[key]
        recs = sorted(g["records"], key=lambda r: (r["round"] if r["round"] else 9,
                                                   r["_order"]))
        used = set()
        main_recs, extra = [], []
        for rec in recs:
            slot = rec["round"] if rec["round"] and rec["round"] not in used else None
            if slot is None:
                slot = next((s for s in (1, 2, 3) if s not in used), None)
            if slot is None:
                extra.append(rec)
            else:
                used.add(slot)
                main_recs.append((slot, rec))

        case = g["case"]
        base_id = case["test_case_id"] if case else key
        if main_recs:
            rounds = {}
            req = []
            for slot, rec in main_recs:
                try:
                    ev = (evaluate_record(case, rec, log, base_id) if case else None)
                except DifyFatalError as e:
                    print(f"fatal: {e}", file=sys.stderr)
                    return 1
                except RuntimeError as e:
                    print(f"[eval {base_id} R{slot}] FAILED: {e}", flush=True)
                    ev = None
                if ev is not None:
                    rounds[slot] = ev
                req.append(slot)
            complete = all(rounds.get(s) for s in req)
            case_overall = (min(rounds[s]["overall"] for s in req)
                            if complete else None)
            rows.append({"row_id": base_id, "case": case, "match": g["match"],
                         "records": [rec for _, rec in main_recs],
                         "rounds": rounds, "required_rounds": req,
                         "case_overall": case_overall})
            n_eval = len(rounds)
            status = ("PASS" if complete and case_overall >= PASS_THRESHOLD else
                      "FAIL" if any(not rd["pass"] for rd in rounds.values()) or
                      (complete and case_overall < PASS_THRESHOLD) else
                      "EVALUATION INCOMPLETE")
            print(f"[test {gi}/{len(groups)}] {base_id}: {status} "
                  f"({n_eval}/{len(req)} rounds evaluated"
                  + (f", min {case_overall:.1%}" if case_overall is not None else "")
                  + ")", flush=True)
        else:
            # golden case with no real answer (or empty unmatched group): stays
            # pending in the denominator, never silently dropped
            rows.append({"row_id": base_id, "case": case, "match": g["match"],
                         "records": [], "rounds": {}, "required_rounds": [],
                         "case_overall": None})
            print(f"[test {gi}/{len(groups)}] {base_id}: "
                  f"Pending (no real system answer)", flush=True)
        for n, rec in enumerate(extra, start=4):
            rows.append({"row_id": f"{base_id}-OBS{n}", "case": case,
                         "match": g["match"], "records": [rec],
                         "rounds": {}, "required_rounds": [],
                         "case_overall": None})
            print(f"[test extra] {base_id}-OBS{n}: skipped (round slots full)",
                  flush=True)

    pdf_name = golden.get("pdf", golden_path.name)
    write_workbook(out_xlsx, template, run_id, golden, rows, pdf_name,
                   run_started, args.review_type)

    log.event("run_finished", rows=len(rows),
              outputs=[str(out_xlsx)])
    log.close()
    print(f"done: {len(rows)} test rows -> {out_xlsx} (log: {log_path})", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
