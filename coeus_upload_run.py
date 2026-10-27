#!/usr/bin/env python3
"""Upload Coeus documents and start parsing for single or batch workflows."""

from __future__ import annotations

import json
import mimetypes
import sys
from pathlib import Path
from typing import Any

import requests

DEFAULT_AUTH_TYPE = "saas"
DEFAULT_TIMEOUT = 60


class CoeusError(RuntimeError):
    """Raised when a Coeus request fails or returns an unsuccessful payload."""


def auth_headers(api_key: str, auth_type: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {api_key}",
        "AuthType": auth_type,
    }


def request_json(
    method: str,
    url: str,
    *,
    headers: dict[str, str],
    timeout: int,
    verify: bool,
    data: dict[str, str] | None = None,
    files: list[tuple[str, tuple[str, Any, str]]] | None = None,
    json_body: dict[str, Any] | None = None,
) -> dict[str, Any]:
    try:
        response = requests.request(
            method,
            url,
            headers=headers,
            data=data,
            files=files,
            json=json_body,
            timeout=timeout,
            verify=verify,
        )
    except requests.RequestException as error:
        raise CoeusError(f"network error: {error}") from error

    if response.status_code >= 400:
        raise CoeusError(f"HTTP {response.status_code}: {response.text[:500]}")

    try:
        payload = response.json()
    except ValueError as error:
        raise CoeusError(f"invalid JSON response: {error}") from error
    if not isinstance(payload, dict):
        raise CoeusError("unexpected response shape: expected a JSON object")

    retcode = payload.get("retcode")
    if retcode not in (0, None):
        raise CoeusError(
            f"Coeus request failed: retcode={retcode} retmsg={payload.get('retmsg')}"
        )
    return payload


def upload_documents(
    base_url: str,
    api_key: str,
    kb_id: str,
    file_paths: list[Path],
    *,
    auth_type: str,
    timeout: int,
    verify: bool,
) -> tuple[dict[str, Any], list[str]]:
    opened_files = []
    try:
        multipart_files = []
        for path in file_paths:
            if not path.is_file():
                raise CoeusError(f"no such file: {path}")
            handle = path.open("rb")
            opened_files.append(handle)
            content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            multipart_files.append(("file", (path.name, handle, content_type)))

        payload = request_json(
            "POST",
            f"{base_url.rstrip('/')}/document/upload",
            headers=auth_headers(api_key, auth_type),
            data={"kb_id": kb_id},
            files=multipart_files,
            timeout=timeout,
            verify=verify,
        )
    finally:
        for handle in opened_files:
            handle.close()

    document_ids = payload.get("data")
    if not isinstance(document_ids, list) or not all(
        isinstance(document_id, str) and document_id for document_id in document_ids
    ):
        raise CoeusError("upload response missing data document ID list")
    return payload, document_ids


def run_documents(
    base_url: str,
    api_key: str,
    document_ids: list[str],
    *,
    auth_type: str,
    timeout: int,
    verify: bool,
) -> dict[str, Any]:
    if not document_ids:
        raise CoeusError("no document IDs returned by upload")
    return request_json(
        "POST",
        f"{base_url.rstrip('/')}/document/run",
        headers=auth_headers(api_key, auth_type),
        json_body={"doc_ids": document_ids, "run": 1},
        timeout=timeout,
        verify=verify,
    )


def upload_and_run(
    base_url: str,
    api_key: str,
    kb_id: str,
    file_paths: list[Path],
    *,
    auth_type: str = DEFAULT_AUTH_TYPE,
    timeout: int = DEFAULT_TIMEOUT,
    verify: bool = True,
) -> dict[str, Any]:
    upload_response, document_ids = upload_documents(
        base_url,
        api_key,
        kb_id,
        file_paths,
        auth_type=auth_type,
        timeout=timeout,
        verify=verify,
    )
    run_response = run_documents(
        base_url,
        api_key,
        document_ids,
        auth_type=auth_type,
        timeout=timeout,
        verify=verify,
    )
    return {
        "files": [str(path) for path in file_paths],
        "document_ids": document_ids,
        "upload": upload_response,
        "run": run_response,
    }


def main() -> int:
    base_url = "https://your-coeus-host"
    api_key = "your-api-key"
    kb_id = "your-kb-id"
    file_paths = [
        Path("/path/to/document.pdf"),
    ]
    # Set to False only when the trusted Coeus endpoint uses a self-signed certificate.
    verify = True

    try:
        result = upload_and_run(
            base_url,
            api_key,
            kb_id,
            file_paths,
            verify=verify,
        )
    except CoeusError as error:
        print(f"error: {error}", file=sys.stderr)
        return 1

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
