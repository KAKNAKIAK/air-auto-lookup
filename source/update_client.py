# -*- coding: utf-8 -*-
import base64
import hashlib
import json
import os
import re
import subprocess
import tempfile
import urllib.request
from pathlib import Path
from urllib.parse import urlparse

DEFAULT_LATEST_URL = "https://raw.githubusercontent.com/KAKNAKIAK/air-auto-lookup/main/latest.json"
DEFAULT_USER_AGENT = "AirAutoLookupUpdateClient/1.0"
MAX_MANIFEST_BYTES = 1024 * 1024


class UpdateError(Exception):
    pass


def parse_version(version: str) -> tuple[int, ...]:
    text = str(version or "").strip().lower()
    text = text[1:] if text.startswith("v") else text
    numbers = [int(part) for part in re.findall(r"\d+", text)]
    return tuple(numbers or [0])


def is_newer_version(latest_version: str, current_version: str) -> bool:
    latest = parse_version(latest_version)
    current = parse_version(current_version)
    size = max(len(latest), len(current))
    latest += (0,) * (size - len(latest))
    current += (0,) * (size - len(current))
    return latest > current


def _read_url(url: str, timeout: float) -> bytes:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": DEFAULT_USER_AGENT,
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        content = response.read(MAX_MANIFEST_BYTES + 1)
    if len(content) > MAX_MANIFEST_BYTES:
        raise UpdateError("latest.json 파일 용량이 초과되었습니다.")
    return content


def _decode_manifest_json(raw: bytes) -> dict[str, object]:
    payload = json.loads(raw.decode("utf-8-sig"))
    if (
        isinstance(payload, dict)
        and payload.get("encoding") == "base64"
        and isinstance(payload.get("content"), str)
    ):
        content = payload["content"].replace("
", "")
        decoded = base64.b64decode(content)
        return json.loads(decoded.decode("utf-8-sig"))
    if isinstance(payload, dict):
        return payload
    raise UpdateError("latest.json 형식이 올바르지 않습니다.")


def fetch_latest_manifest(latest_url: str = DEFAULT_LATEST_URL, timeout: float = 8.0) -> dict[str, str]:
    if not latest_url:
        raise UpdateError("업데이트 확인 URL이 비어 있습니다.")

    if "raw.githubusercontent.com" in latest_url:
        import time
        sep = "&" if "?" in latest_url else "?"
        latest_url = f"{latest_url}{sep}t={int(time.time())}"

    try:
        raw = _read_url(latest_url, timeout)
        manifest = _decode_manifest_json(raw)
    except UpdateError:
        raise
    except Exception as exc:
        raise UpdateError(f"latest.json 확인 실패: {exc}") from exc

    version = str(manifest.get("version", "")).strip()
    download_url = str(
        manifest.get("download_url") or manifest.get("url") or ""
    ).strip()
    sha256 = str(manifest.get("sha256", "")).strip().lower()

    if not version:
        raise UpdateError("latest.json에 version 정보가 없습니다.")
    if not download_url:
        raise UpdateError("latest.json에 download_url 정보가 없습니다.")
    if not re.fullmatch(r"[0-9a-fA-F]{64}", sha256):
        raise UpdateError("latest.json의 sha256 값이 올바르지 않습니다.")

    result: dict[str, str] = {
        "version": version,
        "download_url": download_url,
        "sha256": sha256,
        "release_notes": str(manifest.get("release_notes") or manifest.get("notes") or "").strip(),
    }
    return result


def fetch_available_update(latest_url: str = DEFAULT_LATEST_URL, current_version: str = "v1.0.0", timeout: float = 8.0) -> dict[str, str] | None:
    manifest = fetch_latest_manifest(latest_url, timeout=timeout)
    if is_newer_version(manifest["version"], current_version):
        return manifest
    return None


def download_installer(
    manifest: dict[str, str],
    timeout: float = 30.0,
    progress_callback: object = None,
) -> Path:
    download_url = manifest["download_url"]
    expected_sha256 = manifest["sha256"].lower()
    version = manifest.get("version", "latest")
    file_name = Path(urlparse(download_url).path).name or f"AirAutoLookup_Setup_{version}.exe"
    target_dir = Path(tempfile.mkdtemp(prefix="AirAutoLookupDownload_"))
    target_path = target_dir / file_name

    request = urllib.request.Request(
        download_url,
        headers={"User-Agent": DEFAULT_USER_AGENT},
    )
    digest = hashlib.sha256()
    downloaded = 0

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            total_header = response.headers.get("Content-Length")
            total = int(total_header) if total_header and total_header.isdigit() else None
            with target_path.open("wb") as out_file:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    out_file.write(chunk)
                    digest.update(chunk)
                    downloaded += len(chunk)
                    if callable(progress_callback):
                        progress_callback(downloaded, total)
    except Exception as exc:
        raise UpdateError(f"설치본 다운로드 실패: {exc}") from exc

    actual_sha256 = digest.hexdigest().lower()
    if actual_sha256 != expected_sha256:
        try:
            target_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise UpdateError(
            "설치본 SHA256 검증 실패
"
            f"기대값: {expected_sha256}
"
            f"실제값: {actual_sha256}"
        )

    return target_path


def run_installer_and_exit(setup_exe_path: Path) -> None:
    if not setup_exe_path.exists():
        raise UpdateError(f"설치 파일이 존재하지 않습니다: {setup_exe_path}")
    subprocess.Popen([str(setup_exe_path)], creationflags=getattr(subprocess, "DETACHED_PROCESS", 0))
