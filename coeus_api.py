#!/usr/bin/env python3
"""Call the Coeus Integration API.

Usage:
    python3 coeus_api.py new-conversation --api-key YOUR_API_KEY --dialog-id your-dialog-id
    python3 coeus_api.py upload --api-key YOUR_API_KEY --kb-id YOUR_KB_ID file1.pdf file2.docx
    python3 coeus_api.py run --api-key YOUR_API_KEY DOC_ID1 DOC_ID2
"""
import argparse
import json
import os
import sys
import time
from pathlib import Path

import requests

# TODO: fill in the real Coeus host before use.
BASE_URL = "https://coeus.example.com"

DEFAULT_AUTH_TYPE = "saas"
DEFAULT_TIMEOUT = 30
MAX_RETRIES = 3
RETRY_BACKOFF = 2


class CoeusError(RuntimeError):
    pass


def auth_headers(api_key: str, end_user_token: str | None, auth_type: str) -> dict:
    headers = {"Authorization": f"Bearer {api_key}", "AuthType": auth_type}
    if end_user_token:
        headers["EndUserToken"] = end_user_token
    return headers


def request_json(method, url, *, headers, params=None, data=None, files=None, timeout) -> dict:
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        if attempt > 1:
            time.sleep(RETRY_BACKOFF * (attempt - 1))
        try:
            response = requests.request(
                method,
                url,
                headers=headers,
                params=params,
                data=data,
                files=files,
                timeout=timeout,
            )
        except requests.RequestException as error:
            last_error = f"network error: {error}"
            continue

        if response.status_code >= 500:
            last_error = f"HTTP {response.status_code}: {response.text[:300]}"
            continue
        if response.status_code >= 400:
            raise CoeusError(f"HTTP {response.status_code}: {response.text[:300]}")

        try:
            payload = response.json()
        except ValueError as error:
            raise CoeusError(f"invalid JSON response: {error}") from error
        if not isinstance(payload, dict):
            raise CoeusError("unexpected response shape: expected a JSON object")
        return payload

    raise CoeusError(f"request failed after {MAX_RETRIES} attempts: {last_error}")


def new_conversation(
    api_key: str,
    end_user_token: str | None,
    dialog_id: str | None,
    share_token: str | None,
    auth_type: str,
    timeout: int,
) -> dict:
    url = f"{BASE_URL.rstrip('/')}/api/new_conversation"
    headers = auth_headers(api_key, end_user_token, auth_type)
    if share_token:
        headers["ShareToken"] = share_token
    params = {"dialog_id": dialog_id} if dialog_id else None
    return request_json("GET", url, headers=headers, params=params, timeout=timeout)


def upload_document(
    api_key: str,
    end_user_token: str | None,
    kb_id: str,
    file_paths: list[str],
    auth_type: str,
    timeout: int,
) -> dict:
    url = f"{BASE_URL.rstrip('/')}/document/upload"
    headers = auth_headers(api_key, end_user_token, auth_type)
    data = {"kb_id": kb_id}

    opened_files = []
    try:
        files = []
        for raw_path in file_paths:
            path = Path(raw_path)
            if not path.is_file():
                raise CoeusError(f"no such file: {path}")
            handle = path.open("rb")
            opened_files.append(handle)
            files.append(("file", (path.name, handle)))

        payload = request_json(
            "POST", url, headers=headers, data=data, files=files, timeout=timeout
        )
    finally:
        for handle in opened_files:
            handle.close()

    retcode = payload.get("retcode")
    if retcode not in (0, None):
        raise CoeusError(f"upload failed: retcode={retcode} retmsg={payload.get('retmsg')}")
    return payload


def run_document(
    api_key: str,
    end_user_token: str | None,
    doc_ids: list[str],
    action: str,
    auth_type: str,
    timeout: int,
) -> dict:
    # TODO: /document/run request/response schema is not yet captured in
    # Coeus-API-Spec.md (section 12-13 details missing). Fill in once confirmed.
    raise CoeusError(
        "/document/run request format is not documented yet; "
        "pending confirmation before implementation"
    )


def build_parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument(
        "--api-key",
        required=True,
        help="Coeus API key (sent as 'Authorization: Bearer <api-key>')",
    )
    common.add_argument(
        "--end-user-token",
        default=os.getenv("COEUS_END_USER_TOKEN"),
        help="optional EndUserToken header (defaults to COEUS_END_USER_TOKEN env var)",
    )
    common.add_argument(
        "--auth-type",
        default=os.getenv("COEUS_AUTH_TYPE", DEFAULT_AUTH_TYPE),
        help=f"AuthType header (default: {DEFAULT_AUTH_TYPE})",
    )
    common.add_argument(
        "--timeout",
        type=int,
        default=int(os.getenv("COEUS_TIMEOUT", DEFAULT_TIMEOUT)),
        help=f"request timeout in seconds (default: {DEFAULT_TIMEOUT})",
    )

    parser = argparse.ArgumentParser(description="Call the Coeus Integration API.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    p_new = subparsers.add_parser(
        "new-conversation", parents=[common], help="create a new conversation"
    )
    p_new.add_argument("--dialog-id", help="dialog ID; use --share-token when omitted")
    p_new.add_argument("--share-token", help="optional ShareToken header")

    p_upload = subparsers.add_parser(
        "upload", parents=[common], help="upload one or more files to a knowledge base"
    )
    p_upload.add_argument("files", nargs="+", help="local file path(s) to upload")
    p_upload.add_argument("--kb-id", required=True, help="target knowledge base ID")

    p_run = subparsers.add_parser(
        "run",
        parents=[common],
        help="start or cancel parsing for uploaded document(s) (not yet implemented)",
    )
    p_run.add_argument("doc_ids", nargs="+", help="document ID(s) returned by upload")
    p_run.add_argument(
        "--action",
        choices=["run", "cancel"],
        default="run",
        help="start or cancel parsing (default: run)",
    )

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        if args.command == "new-conversation":
            if not args.dialog_id and not args.share_token:
                parser.error("one of --dialog-id or --share-token is required")
            response = new_conversation(
                api_key=args.api_key,
                end_user_token=args.end_user_token,
                dialog_id=args.dialog_id,
                share_token=args.share_token,
                auth_type=args.auth_type,
                timeout=args.timeout,
            )
        elif args.command == "upload":
            response = upload_document(
                api_key=args.api_key,
                end_user_token=args.end_user_token,
                kb_id=args.kb_id,
                file_paths=args.files,
                auth_type=args.auth_type,
                timeout=args.timeout,
            )
        elif args.command == "run":
            response = run_document(
                api_key=args.api_key,
                end_user_token=args.end_user_token,
                doc_ids=args.doc_ids,
                action=args.action,
                auth_type=args.auth_type,
                timeout=args.timeout,
            )
        else:
            parser.error(f"unknown command: {args.command}")
            return 2
    except CoeusError as error:
        print(f"error: {error}", file=sys.stderr)
        return 1

    print(json.dumps(response, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
