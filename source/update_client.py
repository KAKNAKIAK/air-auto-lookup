# -*- coding: utf-8 -*-
"""GitHub Release 기반 항공자동조회 업데이트 도우미."""

from __future__ import annotations

import hashlib
import json
import re
import base64
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path
from urllib.parse import urlparse


DEFAULT_LATEST_URL = "https://api.github.com/repos/KAKNAKIAK/air-auto-lookup/contents/latest.json?ref=main"
DEFAULT_USER_AGENT = "AirAutoLookupUpdateClient/1.0"
MAX_MANIFEST_BYTES = 1024 * 1024


class UpdateError(Exception):
    """업데이트 확인 또는 설치 실패."""


def parse_version(value: str) -> tuple[int, ...]:
    text = str(value or "").strip().lower().lstrip("v")
    return tuple(int(part) for part in re.findall(r"\d+", text)) or (0,)


def is_newer_version(latest_version: str, current_version: str) -> bool:
    latest = parse_version(latest_version)
    current = parse_version(current_version)
    length = max(len(latest), len(current))
    return latest + (0,) * (length - len(latest)) > current + (0,) * (length - len(current))


def _read_url(url: str, timeout: float) -> bytes:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": DEFAULT_USER_AGENT, "Accept": "application/json"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        content = response.read(MAX_MANIFEST_BYTES + 1)
    if len(content) > MAX_MANIFEST_BYTES:
        raise UpdateError("업데이트 정보 파일 용량이 너무 큽니다.")
    return content


def fetch_latest_manifest(latest_url: str = DEFAULT_LATEST_URL, timeout: float = 8.0) -> dict[str, str]:
    if not latest_url:
        raise UpdateError("업데이트 확인 URL이 비어 있습니다.")

    separator = "&" if "?" in latest_url else "?"
    request_url = f"{latest_url}{separator}t={int(time.time())}"
    try:
        payload = json.loads(_read_url(request_url, timeout).decode("utf-8-sig"))
    except Exception as exc:
        raise UpdateError(f"업데이트 정보를 확인하지 못했습니다: {exc}") from exc

    if isinstance(payload, dict) and payload.get("encoding") == "base64" and isinstance(payload.get("content"), str):
        decoded = base64.b64decode(payload["content"].replace("\n", ""))
        payload = json.loads(decoded.decode("utf-8-sig"))
    if not isinstance(payload, dict):
        raise UpdateError("업데이트 정보 형식이 올바르지 않습니다.")
    version = str(payload.get("version") or "").strip()
    download_url = str(payload.get("download_url") or payload.get("url") or "").strip()
    sha256 = str(payload.get("sha256") or "").strip().lower()
    if not version or not download_url or not re.fullmatch(r"[0-9a-f]{64}", sha256):
        raise UpdateError("업데이트 정보에 버전·다운로드 주소·해시가 올바르게 입력되지 않았습니다.")
    return {
        "version": version,
        "download_url": download_url,
        "sha256": sha256,
        "release_notes": str(payload.get("release_notes") or payload.get("notes") or "").strip(),
    }


def fetch_available_update(
    latest_url: str = DEFAULT_LATEST_URL,
    current_version: str = "v1.0.0",
    timeout: float = 8.0,
) -> dict[str, str] | None:
    manifest = fetch_latest_manifest(latest_url, timeout)
    return manifest if is_newer_version(manifest["version"], current_version) else None


def download_installer(manifest: dict[str, str], timeout: float = 60.0) -> Path:
    target_dir = Path(tempfile.mkdtemp(prefix="AirAutoLookupUpdate_"))
    filename = Path(urlparse(manifest["download_url"]).path).name or "AirAutoLookup_Setup.exe"
    target_path = target_dir / filename
    digest = hashlib.sha256()
    request = urllib.request.Request(manifest["download_url"], headers={"User-Agent": DEFAULT_USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response, target_path.open("wb") as output:
            while chunk := response.read(1024 * 1024):
                output.write(chunk)
                digest.update(chunk)
    except Exception as exc:
        raise UpdateError(f"업데이트 설치 파일을 다운로드하지 못했습니다: {exc}") from exc
    if digest.hexdigest().lower() != manifest["sha256"].lower():
        target_path.unlink(missing_ok=True)
        raise UpdateError("다운로드한 설치 파일의 SHA256 검증에 실패했습니다.")
    return target_path


def run_installer_and_exit(installer_path: Path) -> None:
    if not installer_path.exists():
        raise UpdateError(f"설치 파일을 찾지 못했습니다: {installer_path}")
    subprocess.Popen([str(installer_path)], creationflags=getattr(subprocess, "DETACHED_PROCESS", 0))
