from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta
import json
import os
from pathlib import Path
import re
import sqlite3
import time
import urllib.request


TOPAS_PROMPT_INPUT = "#cryptics1_cmd_shellbridge_shellWindow_top_left_modeString_cmdPromptInput"
TOPAS_SHELL_ROOT = "#cryptics1_cmd_shellbridge_shellWindow_top_left"
PROMPT_READY_RE = re.compile(r"(?m)^\s*>\s*$")
LOADING_RE = re.compile(r"(\|\||loading|processing|조회중)", re.IGNORECASE)
TOPAS_COMMAND_RE = re.compile(
    r"AN(?P<date>\d{1,2}[A-Z]{3})(?P<origin>[A-Z]{3})(?P<destination>[A-Z]{3})/A",
    re.IGNORECASE,
)
TOPAS_AVAILABILITY_HEADER_RE = re.compile(
    r"(?m)^\s*\*\*\s+AMADEUS AVAILABILITY\s+-\s+AN\s+\*\*\s+"
    r"(?P<destination>[A-Z]{3})\b.*?\b(?P<date>\d{1,2}[A-Z]{3})\b",
    re.IGNORECASE,
)
RETRY_STATE_FILE = "retry-state.json"
RAW_STORE_FILENAME = "raw-store.sqlite"
DEFAULT_MAX_COLLECT_ERROR_RETRIES = 3
DEFAULT_SESSION_ERROR_THRESHOLD = 3
TOPAS_AC1_GROUP_SIZE = 90
TOPAS_AC1_SEND_CHUNK_SIZE = 10
TOPAS_AC1_KEY_INTERVAL_SECONDS = 0.12
TOPAS_AC1_INPUT_RETRIES = 3
TOPAS_AC1_SINGLE_RESPONSE_TIMEOUT = 15.0
TOPAS_AC1_BLANK_RETRY_LIMIT = 2
TOPAS_AC1_BLANK_SETTLE_SECONDS = 0.7
TOPAS_AC1_MISSING_RETRY_IDLE_SECONDS = 8.0
TOPAS_AC1_MISSING_RETRY_LIMIT = 2
TOPAS_AC1_MISSING_RETRY_MAX_COMMANDS = 2
DEFAULT_COLLECT_BATCH_SIZE = TOPAS_AC1_GROUP_SIZE + 1
RAW_WRITE_WORKERS = 8
RAW_COMPLETE_MARKERS = ("AMADEUS AVAILABILITY", "NO FLIGHT", "REQUEST NEW AVAILABILITY")
MONTH_NAMES = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
PAUSE_CONTROL_FILE = "pause.flag"
STOP_CONTROL_FILE = "stop.flag"
TOPAS_AC1_REQUEST_RE = re.compile(
    r"(?m)^\s*AN(?P<date>\d{1,2}[A-Z]{3})(?P<origin>[A-Z]{3})(?P<destination>[A-Z]{3})"
    r"/A(?P<airline>[A-Z0-9]{2})(?P<flight>\d{0,4})\s+-AC-",
    re.IGNORECASE,
)
TOPAS_FULL_COMMAND_RE = re.compile(
    r"AN(?P<date>\d{1,2}[A-Z]{3})(?P<origin>[A-Z]{3})(?P<destination>[A-Z]{3})"
    r"/A(?P<airline>[A-Z0-9]{2})(?P<flight>\d{0,4})",
    re.IGNORECASE,
)


class TopasSessionError(RuntimeError):
    pass


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def make_event(event_type: str, **payload: object) -> dict[str, object]:
    return {"timestamp": now_iso(), "event": event_type, **payload}


def append_event(log_dir: Path, event_type: str, **payload: object) -> None:
    append_events(log_dir, [make_event(event_type, **payload)])


def append_events(log_dir: Path, events: list[dict[str, object]]) -> None:
    if not events:
        return
    with (log_dir / "events.jsonl").open("a", encoding="utf-8") as handle:
        for event in events:
            handle.write(json.dumps(event, ensure_ascii=False) + "\n")


def load_plan(log_dir: Path, direction: str) -> list[dict[str, object]]:
    plan_path = log_dir / "command-plan.json"
    if not plan_path.exists():
        raise FileNotFoundError(f"command-plan.json을 찾지 못했습니다: {plan_path}")
    plan = json.loads(plan_path.read_text(encoding="utf-8"))

    departure_rows = [row for row in plan.get("departureCommands", []) if isinstance(row, dict)]
    return_rows = [row for row in plan.get("returnCandidateCommands", []) if isinstance(row, dict)]
    if direction == "departure":
        return departure_rows
    if direction == "return":
        return return_rows
    if direction == "all":
        return interleave_commands_by_route(departure_rows, return_rows)
    return []


def interleave_commands_by_route(
    departure_rows: list[dict[str, object]],
    return_rows: list[dict[str, object]],
) -> list[dict[str, object]]:
    route_order: list[str] = []
    grouped: dict[str, dict[str, list[dict[str, object]]]] = {}
    for direction, rows in (("departure", departure_rows), ("return", return_rows)):
        for row in rows:
            route_key = command_route_order_key(row)
            if route_key not in grouped:
                grouped[route_key] = {"departure": [], "return": []}
                route_order.append(route_key)
            grouped[route_key][direction].append(row)

    ordered: list[dict[str, object]] = []
    for route_key in route_order:
        ordered.extend(grouped[route_key]["departure"])
        ordered.extend(grouped[route_key]["return"])
    return ordered


def command_route_order_key(row: dict[str, object]) -> str:
    return (
        str(row.get("routeKey") or "").strip()
        or str(row.get("rawKey") or "").strip()
        or str(row.get("route") or "").strip()
        or "UNKNOWN_ROUTE"
    )


def load_rows(path: Path, label: str) -> list[dict[str, object]]:
    if not path.exists():
        raise FileNotFoundError(f"{label} 파일을 찾지 못했습니다: {path}")
    payload = json.loads(path.read_text(encoding="utf-8"))
    rows = payload.get("rows", []) if isinstance(payload, dict) else payload
    return [row for row in rows if isinstance(row, dict)] if isinstance(rows, list) else []


def command_key(row: dict[str, object]) -> tuple[str, str]:
    return str(row.get("routeKey", "")), str(row.get("baseDepartureDate", ""))


def command_query_date(row: dict[str, object]) -> date | None:
    value = str(row.get("queryDate") or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return None


def topas_day_month(value: date) -> str:
    return f"{value.day:02d}{MONTH_NAMES[value.month - 1]}"


def normalize_topas_flight_number(value: object, airline: object = "") -> str:
    text = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    airline_text = re.sub(r"[^A-Z0-9]", "", str(airline or "").upper())
    if airline_text and text.startswith(airline_text):
        text = text[len(airline_text) :]
    digits = re.sub(r"\D", "", text)
    return str(int(digits)) if digits else ""


def command_signature_from_match(match: re.Match[str]) -> tuple[str, str, str, str, str]:
    airline = match.group("airline").upper()
    return (
        match.group("date").upper().zfill(5),
        match.group("origin").upper(),
        match.group("destination").upper(),
        airline,
        normalize_topas_flight_number(match.group("flight"), airline),
    )


def command_signature_from_text(value: object) -> tuple[str, str, str, str, str] | None:
    match = TOPAS_FULL_COMMAND_RE.search(str(value or "").upper())
    return command_signature_from_match(match) if match else None


def command_signature_from_row(row: dict[str, object]) -> tuple[str, str, str, str, str] | None:
    command_signature = command_signature_from_text(row.get("command"))
    if command_signature:
        return command_signature
    query_date = command_query_date(row)
    airline = str(row.get("airline") or "").strip().upper()
    if query_date is None or not airline:
        return None
    return (
        topas_day_month(query_date).upper().zfill(5),
        str(row.get("origin") or "").strip().upper(),
        str(row.get("destination") or "").strip().upper(),
        airline,
        normalize_topas_flight_number(row.get("flight"), airline),
    )


def first_response_command_signature(text: str) -> tuple[str, str, str, str, str] | None:
    for line in str(text or "").replace("\r\n", "\n").splitlines():
        if not line.strip():
            continue
        return command_signature_from_text(line)
    return command_signature_from_text(text)


def command_response_mismatch(row: dict[str, object], raw_text: str) -> bool:
    expected = command_signature_from_row(row)
    actual = first_response_command_signature(raw_text)
    return bool(expected and actual and expected != actual)


def raw_text_state_for_command(row: dict[str, object], raw_text: str) -> str:
    if command_response_mismatch(row, raw_text):
        return "command_mismatch"
    return raw_text_state(raw_text)


def command_identity(row: dict[str, object]) -> str:
    parts = [
        str(row.get("direction") or ""),
        str(row.get("command") or ""),
        str(row.get("queryDate") or ""),
        str(row.get("origin") or ""),
        str(row.get("destination") or ""),
        str(row.get("airline") or ""),
        str(row.get("flight") or ""),
    ]
    identity = "|".join(part.strip().upper() for part in parts if part.strip())
    return identity or str(row.get("rawFile") or row.get("id") or "")


def command_raw_files(row: dict[str, object]) -> list[Path]:
    raw_files: list[Path] = []
    raw_file_values = row.get("rawFiles")
    if isinstance(raw_file_values, list):
        raw_files.extend(Path(str(value)) for value in raw_file_values if str(value or "").strip())
    raw_file = str(row.get("rawFile") or "").strip()
    if raw_file:
        raw_files.append(Path(raw_file))

    selected: list[Path] = []
    seen: set[str] = set()
    for raw_path in raw_files:
        key = str(raw_path)
        if key in seen:
            continue
        seen.add(key)
        selected.append(raw_path)
    return selected


def raw_text_state(text: str) -> str:
    stripped = str(text or "").strip()
    if not stripped:
        return "empty_or_truncated"
    upper = stripped.upper()
    first_line = upper.splitlines()[0] if upper.splitlines() else ""
    if first_line.startswith("TOPAS_COLLECT_ERROR") or first_line.startswith("ERROR"):
        return "collect_error"
    if "NO FLIGHT" in upper:
        return "no_flight"
    if any(marker in upper for marker in RAW_COMPLETE_MARKERS):
        return "normal_raw"
    return "empty_or_truncated"


def raw_store_path(log_dir: Path) -> Path:
    run_path = log_dir / "run.json"
    if run_path.exists():
        try:
            run_doc = json.loads(run_path.read_text(encoding="utf-8"))
        except Exception:
            run_doc = {}
        if isinstance(run_doc, dict):
            raw_store = str(run_doc.get("rawStore") or "").strip()
            if raw_store:
                return Path(raw_store)
            raw_root = str(run_doc.get("rawRoot") or "").strip()
            if raw_root:
                return Path(raw_root) / RAW_STORE_FILENAME
    return log_dir / RAW_STORE_FILENAME


def ensure_raw_store(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS raw_responses (
            command_identity TEXT PRIMARY KEY,
            primary_raw_file TEXT NOT NULL,
            raw_files_json TEXT NOT NULL,
            route_key TEXT NOT NULL,
            raw_key TEXT NOT NULL,
            route TEXT NOT NULL,
            direction TEXT NOT NULL,
            base_departure_date TEXT NOT NULL,
            query_date TEXT NOT NULL,
            candidate_nights TEXT NOT NULL,
            candidate_return_date TEXT NOT NULL,
            command TEXT NOT NULL,
            origin TEXT NOT NULL,
            destination TEXT NOT NULL,
            airline TEXT NOT NULL,
            flight TEXT NOT NULL,
            raw_text TEXT NOT NULL,
            status TEXT NOT NULL,
            error TEXT NOT NULL,
            collection_mode TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_raw_responses_status ON raw_responses(status)")
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_raw_responses_route_date "
        "ON raw_responses(route_key, direction, query_date)"
    )


def open_raw_store(log_dir: Path) -> sqlite3.Connection:
    path = raw_store_path(log_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    ensure_raw_store(conn)
    return conn


def raw_record_for_command(log_dir: Path, row: dict[str, object]) -> dict[str, object] | None:
    path = raw_store_path(log_dir)
    if not path.exists():
        return None
    try:
        conn = sqlite3.connect(path)
        conn.row_factory = sqlite3.Row
        try:
            found = conn.execute(
                "SELECT raw_text, status, error FROM raw_responses WHERE command_identity = ?",
                (command_identity(row),),
            ).fetchone()
        finally:
            conn.close()
    except Exception:
        return None
    return dict(found) if found is not None else None


def load_raw_store_records(log_dir: Path) -> dict[str, dict[str, object]]:
    path = raw_store_path(log_dir)
    if not path.exists():
        return {}
    try:
        conn = sqlite3.connect(path)
        conn.row_factory = sqlite3.Row
        try:
            rows = conn.execute("SELECT command_identity, raw_text, status, error FROM raw_responses").fetchall()
        finally:
            conn.close()
    except Exception:
        return {}
    return {str(row["command_identity"]): dict(row) for row in rows}


def raw_record_from_cache(
    raw_records: dict[str, dict[str, object]] | None,
    row: dict[str, object],
) -> dict[str, object] | None:
    if not raw_records:
        return None
    return raw_records.get(command_identity(row))


def raw_response_state_from_record(
    record: dict[str, object] | None,
    raw_path: Path,
    row: dict[str, object] | None = None,
) -> str:
    if record is not None:
        if row is not None and command_response_mismatch(row, str(record.get("raw_text") or "")):
            return "command_mismatch"
        status = str(record.get("status") or "").strip()
        return status or raw_text_state(str(record.get("raw_text") or ""))
    return raw_file_state(raw_path)


def retry_state_path(log_dir: Path) -> Path:
    return log_dir / RETRY_STATE_FILE



def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        with temp_path.open("w", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass

def load_retry_state(log_dir: Path) -> dict[str, object]:
    path = retry_state_path(log_dir)
    if not path.exists():
        return {"version": 1, "commands": {}}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"version": 1, "commands": {}}
    if not isinstance(payload, dict):
        return {"version": 1, "commands": {}}
    commands = payload.get("commands")
    if not isinstance(commands, dict):
        payload["commands"] = {}
    payload.setdefault("version", 1)
    return payload


def save_retry_state(log_dir: Path, retry_state: dict[str, object]) -> None:
    atomic_write_text(retry_state_path(log_dir), json.dumps(retry_state, ensure_ascii=False, indent=2) + "\n")

def retry_record(retry_state: dict[str, object], identity: str) -> dict[str, object] | None:
    commands = retry_state.get("commands", {})
    if not isinstance(commands, dict):
        return None
    record = commands.get(identity)
    return record if isinstance(record, dict) else None


def retry_count(retry_state: dict[str, object], row: dict[str, object]) -> int:
    record = retry_record(retry_state, command_identity(row))
    if not record:
        return 0
    try:
        return int(record.get("retryCount") or 0)
    except (TypeError, ValueError):
        return 0


def update_retry_failure(
    log_dir: Path,
    retry_state: dict[str, object],
    row: dict[str, object],
    error: str,
    max_retries: int,
) -> dict[str, object]:
    identity = command_identity(row)
    commands = retry_state.setdefault("commands", {})
    if not isinstance(commands, dict):
        commands = {}
        retry_state["commands"] = commands
    now = now_iso()
    record = commands.get(identity)
    if not isinstance(record, dict):
        record = {
            "retryCount": 0,
            "firstFailedAt": now,
        }
    record["retryCount"] = int(record.get("retryCount") or 0) + 1
    record.setdefault("firstFailedAt", now)
    record["lastTriedAt"] = now
    record["lastError"] = error
    record["topasCommand"] = str(row.get("command") or "")
    record["direction"] = str(row.get("direction") or "")
    record["rawFiles"] = [str(path) for path in command_raw_files(row)]
    record["retryExhausted"] = int(record["retryCount"]) >= max_retries
    commands[identity] = record
    save_retry_state(log_dir, retry_state)
    return record


def clear_retry_record(log_dir: Path, retry_state: dict[str, object], row: dict[str, object]) -> None:
    clear_retry_records(log_dir, retry_state, [row])


def clear_retry_records(log_dir: Path, retry_state: dict[str, object], rows: list[dict[str, object]]) -> None:
    commands = retry_state.get("commands", {})
    changed = False
    if isinstance(commands, dict):
        for row in rows:
            identity = command_identity(row)
            if identity in commands:
                commands.pop(identity, None)
                changed = True
    if changed:
        save_retry_state(log_dir, retry_state)


def raw_file_state(path: Path) -> str:
    if not path.exists():
        return "missing"
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return "collect_error"
    stripped = text.strip()
    if not stripped:
        return "empty_or_truncated"
    upper = stripped.upper()
    first_line = upper.splitlines()[0] if upper.splitlines() else ""
    if first_line.startswith("TOPAS_COLLECT_ERROR") or first_line.startswith("ERROR"):
        return "collect_error"
    if "NO FLIGHT" in upper:
        return "no_flight"
    if any(marker in upper for marker in RAW_COMPLETE_MARKERS):
        return "normal_raw"
    return "empty_or_truncated"


def raw_response_state(log_dir: Path, row: dict[str, object], raw_path: Path) -> str:
    record = raw_record_for_command(log_dir, row)
    return raw_response_state_from_record(record, raw_path, row)


def command_has_actionable_raw(
    log_dir: Path,
    row: dict[str, object],
    retry_state: dict[str, object],
    retry_collect_errors: bool,
    max_retries: int,
    raw_records: dict[str, dict[str, object]] | None = None,
) -> bool:
    raw_files = command_raw_files(row)
    if not raw_files:
        return False
    count = retry_count(retry_state, row)
    raw_record = raw_record_from_cache(raw_records, row)
    for raw_path in raw_files:
        state = (
            raw_response_state_from_record(raw_record, raw_path, row)
            if raw_records is not None
            else raw_response_state(log_dir, row, raw_path)
        )
        if state == "missing":
            return True
        if retry_collect_errors and state in {"collect_error", "empty_or_truncated", "command_mismatch"} and count < max_retries:
            return True
    return False


def collect_pending_summary(log_dir: Path, max_retries: int = DEFAULT_MAX_COLLECT_ERROR_RETRIES) -> dict[str, object]:
    rows = load_plan(log_dir, "all")
    merged = merge_duplicate_commands(rows)
    retry_state = load_retry_state(log_dir)
    raw_records = load_raw_store_records(log_dir)
    unique_raw_files: set[str] = set()
    counts = {
        "missing": 0,
        "normal_raw": 0,
        "no_flight": 0,
        "collect_error_retryable": 0,
        "empty_or_truncated_retryable": 0,
        "command_mismatch_retryable": 0,
        "retry_exhausted": 0,
    }
    actionable = 0
    for row in merged:
        identity_retry_count = retry_count(retry_state, row)
        command_actionable = False
        raw_record = raw_record_from_cache(raw_records, row)
        for raw_path in command_raw_files(row):
            unique_raw_files.add(str(raw_path))
            state = raw_response_state_from_record(raw_record, raw_path, row)
            if state == "missing":
                counts["missing"] += 1
                command_actionable = True
            elif state in {"collect_error", "empty_or_truncated", "command_mismatch"}:
                if identity_retry_count < max_retries:
                    key = f"{state}_retryable"
                    counts[key] += 1
                    command_actionable = True
                else:
                    counts["retry_exhausted"] += 1
            else:
                counts[state] += 1
        if command_actionable:
            actionable += 1
    return {
        "logicalRows": len(rows),
        "uniqueTopasCommands": len(merged),
        "uniqueRawFiles": len(unique_raw_files),
        "actionablePendingCommands": actionable,
        **counts,
    }


def actionable_pending_count(pending: dict[str, object]) -> int:
    try:
        return int(pending.get("actionablePendingCommands") or 0)
    except (TypeError, ValueError):
        return 0


def merge_duplicate_commands(commands: list[dict[str, object]]) -> list[dict[str, object]]:
    merged: dict[str, dict[str, object]] = {}
    ordered: list[dict[str, object]] = []
    for row in commands:
        identity = command_identity(row)
        raw_files = [str(path) for path in command_raw_files(row)]
        if identity in merged:
            existing = merged[identity]
            existing_raw_files = list(existing.get("rawFiles", [])) if isinstance(existing.get("rawFiles"), list) else []
            for raw_file in raw_files:
                if raw_file not in existing_raw_files:
                    existing_raw_files.append(raw_file)
            existing["rawFiles"] = existing_raw_files
            continue
        merged_row = dict(row)
        merged_row["rawFiles"] = raw_files
        if raw_files:
            merged_row["rawFile"] = raw_files[0]
        merged[identity] = merged_row
        ordered.append(merged_row)
    return ordered


def is_absolute_an_command(row: dict[str, object]) -> bool:
    return bool(TOPAS_COMMAND_RE.search(str(row.get("command") or "").strip()))


def ac1_group_key(row: dict[str, object]) -> tuple[str, str, str, str, str, str]:
    return (
        str(row.get("direction") or "").strip().lower(),
        str(row.get("origin") or "").strip().upper(),
        str(row.get("destination") or "").strip().upper(),
        str(row.get("airline") or "").strip().upper(),
        str(row.get("flight") or "").strip().upper(),
        str(row.get("routeKey") or "").strip(),
    )


def can_follow_with_ac1(previous: dict[str, object], current: dict[str, object]) -> bool:
    previous_date = command_query_date(previous)
    current_date = command_query_date(current)
    if previous_date is None or current_date is None:
        return False
    return (
        ac1_group_key(previous) == ac1_group_key(current)
        and is_absolute_an_command(previous)
        and is_absolute_an_command(current)
        and current_date == previous_date + timedelta(days=1)
    )


def build_ac1_collection_groups(
    commands: list[dict[str, object]],
    ac1_batch_size: int = TOPAS_AC1_GROUP_SIZE,
) -> list[list[dict[str, object]]]:
    groups: list[list[dict[str, object]]] = []
    current: list[dict[str, object]] = []
    max_group_size = max(1, ac1_batch_size) + 1
    for row in commands:
        if not current:
            current = [row]
            continue
        if len(current) < max_group_size and can_follow_with_ac1(current[-1], row):
            current.append(row)
            continue
        groups.append(current)
        current = [row]
    if current:
        groups.append(current)
    return groups


def filter_commands(
    log_dir: Path,
    commands: list[dict[str, object]],
    mode: str,
    retry_state: dict[str, object] | None = None,
    retry_collect_errors: bool = False,
    max_retries: int = DEFAULT_MAX_COLLECT_ERROR_RETRIES,
    raw_records: dict[str, dict[str, object]] | None = None,
) -> list[dict[str, object]]:
    if mode == "all":
        return commands
    if mode == "raw-missing":
        state = retry_state or load_retry_state(log_dir)
        records = raw_records if raw_records is not None else load_raw_store_records(log_dir)
        return [
            row
            for row in commands
            if command_has_actionable_raw(log_dir, row, state, retry_collect_errors, max_retries, records)
        ]
    if mode == "unconfirmed":
        rows = load_rows(log_dir / "return-night-results.json", "return-night-results.json")
        targets = {
            (str(row.get("routeKey", "")), str(row.get("baseDepartureDate", "")))
            for row in rows
            if str(row.get("status", "")) != "confirmed"
        }
        return [row for row in commands if command_key(row) in targets]
    if mode == "fare-route-missing":
        rows = load_rows(log_dir / "fare-results.json", "fare-results.json")
        targets = {
            (str(row.get("routeKey", "")), str(row.get("baseDepartureDate", "")))
            for row in rows
            if str(row.get("status", "")) == "fare_route_missing"
        }
        return [row for row in commands if command_key(row) in targets]
    raise ValueError(f"지원하지 않는 필터입니다: {mode}")


def is_session_error(exc: Exception | str) -> bool:
    text = str(exc).lower()
    session_markers = (
        "디버그 브라우저 연결 실패",
        "topas entry 탭을 찾지 못했습니다",
        "topas shell 영역을 찾지 못했습니다",
        "topas 명령 입력창을 찾지 못했습니다",
        "chrome not reachable",
        "cannot connect",
        "no such window",
        "invalid session id",
        "command prompt",
        "prompt not found",
        "input not found",
        "element not interactable",
        "login",
        "로그인",
    )
    return any(marker in text for marker in session_markers)


def update_run_status(log_dir: Path, status: str, **updates: object) -> None:
    run_path = log_dir / "run.json"
    if not run_path.exists():
        return
    try:
        run_doc = json.loads(run_path.read_text(encoding="utf-8"))
    except Exception:
        return
    run_doc["status"] = status
    run_doc["statusUpdatedAt"] = now_iso()
    run_doc.update(updates)
    atomic_write_text(run_path, json.dumps(run_doc, ensure_ascii=False, indent=2) + "\n")



def read_run_status(log_dir: Path) -> str:
    run_path = log_dir / "run.json"
    if not run_path.exists():
        return ""
    try:
        run_doc = json.loads(run_path.read_text(encoding="utf-8"))
    except Exception:
        return ""
    return str(run_doc.get("status") or "")


def debugger_targets(address: str, timeout: float = 1.5) -> list[dict[str, object]]:
    url = f"http://{address}/json/list"
    with urllib.request.urlopen(url, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8", errors="replace"))
    return payload if isinstance(payload, list) else []


def target_looks_like_topas(target: dict[str, object]) -> bool:
    title = str(target.get("title") or "").lower()
    url = str(target.get("url") or "").lower()
    return "topassellconnect.com" in url or "topas" in url or "sell connect" in title


def require_topas_debug_target(address: str) -> None:
    targets = debugger_targets(address)
    if not any(target_looks_like_topas(target) for target in targets if isinstance(target, dict)):
        raise RuntimeError(f"{address} 디버그 브라우저에서 TOPAS 탭을 찾지 못했습니다.")

def control_path(log_dir: Path, filename: str) -> Path:
    return log_dir / filename


def pause_requested(log_dir: Path) -> bool:
    return control_path(log_dir, PAUSE_CONTROL_FILE).exists()


def stop_requested(log_dir: Path) -> bool:
    return control_path(log_dir, STOP_CONTROL_FILE).exists()


def wait_if_paused(log_dir: Path) -> bool:
    if not pause_requested(log_dir):
        return False
    append_event(log_dir, "topas_collection_paused")
    update_run_status(log_dir, "paused", controlState="paused")
    print("일시중지 요청 확인: 현재 상태로 대기합니다. 재개 또는 정지를 눌러 주세요.")
    while pause_requested(log_dir) and not stop_requested(log_dir):
        time.sleep(1)
    if stop_requested(log_dir):
        append_event(log_dir, "topas_collection_stop_requested", phase="paused")
        return True
    append_event(log_dir, "topas_collection_resumed")
    update_run_status(log_dir, "running", controlState="running")
    print("재개 요청 확인: 수집을 이어갑니다.")
    return False


class TopasLiveCollector:
    def __init__(self, debugger_addresses: list[str], timeout: float = 80.0):
        self.debugger_addresses = debugger_addresses
        self.timeout = timeout
        self.window_handle = None
        self.driver = None
        self.debugger_address = ""

    def ensure_connected(self, log_dir: Path):
        if self.driver is not None:
            try:
                self.switch_to_topas(self.driver)
                return self.driver, self.debugger_address
            except Exception as exc:
                append_event(log_dir, "topas_session_error", error=str(exc), phase="reuse")
                self.close()
        try:
            self.driver, self.debugger_address = self.connect()
            append_event(log_dir, "topas_live_connected", debuggerAddress=self.debugger_address)
            return self.driver, self.debugger_address
        except Exception as exc:
            append_event(log_dir, "topas_session_error", error=str(exc), phase="connect")
            update_run_status(log_dir, "session_error", lastError=str(exc))
            raise TopasSessionError(str(exc)) from exc

    def close(self) -> None:
        if self.driver is None:
            return
        try:
            self.driver.quit()
        except Exception:
            pass
        finally:
            self.driver = None
            self.debugger_address = ""
            self.window_handle = None

    def collect(
        self,
        commands: list[dict[str, object]],
        log_dir: Path,
        continue_on_error: bool = False,
        retries: int = 0,
        write_error_raw: bool = False,
        max_collect_error_retries: int = DEFAULT_MAX_COLLECT_ERROR_RETRIES,
        session_error_threshold: int = DEFAULT_SESSION_ERROR_THRESHOLD,
    ) -> list[dict[str, object]]:
        driver, address = self.ensure_connected(log_dir)
        results: list[dict[str, object]] = []
        retry_state = load_retry_state(log_dir)
        consecutive_session_errors = 0
        try:
            command_index = 0
            for group in build_ac1_collection_groups(commands):
                if len(group) > 1:
                    try:
                        group_results = self.collect_ac1_group(driver, group, log_dir, retry_state, command_index + 1)
                        results.extend(group_results)
                        command_index += len(group_results)
                        consecutive_session_errors = 0
                        continue
                    except Exception as exc:
                        append_event(
                            log_dir,
                            "topas_ac1_batch_fallback",
                            commandCount=len(group),
                            error=str(exc),
                            sessionError=is_session_error(exc),
                        )
                        print(f"AC1 묶음 수집 실패, 기존 AN 개별 조회로 전환합니다: {str(exc).splitlines()[0]}")
                else:
                    group = group[:1]

                for command_row in group:
                    command_index += 1
                    result = self.collect_single_command(
                        driver,
                        command_row,
                        log_dir,
                        retry_state,
                        command_index,
                        continue_on_error,
                        retries,
                        write_error_raw,
                        max_collect_error_retries,
                    )
                    results.append(result)
                    if result.get("status") == "session_error":
                        consecutive_session_errors += 1
                        if consecutive_session_errors >= session_error_threshold:
                            message = f"TOPAS 세션 오류가 {consecutive_session_errors}회 연속 발생해 수집을 중단합니다."
                            append_event(
                                log_dir,
                                "topas_session_error_circuit_opened",
                                threshold=session_error_threshold,
                                error=message,
                            )
                            update_run_status(log_dir, "session_error", lastError=message)
                            raise TopasSessionError(message)
                    else:
                        consecutive_session_errors = 0
        finally:
            pass
        return results

    def collect_single_command(
        self,
        driver,
        command_row: dict[str, object],
        log_dir: Path,
        retry_state: dict[str, object],
        index: int,
        continue_on_error: bool,
        retries: int,
        write_error_raw: bool,
        max_collect_error_retries: int,
    ) -> dict[str, object]:
        command = str(command_row.get("command") or "").strip()
        raw_files = command_raw_files(command_row)
        raw_file = raw_files[0] if raw_files else Path("")
        started = time.perf_counter()
        for attempt in range(retries + 1):
            try:
                raw_text = self.execute_command(driver, command)
                self.write_raw_response(log_dir, command_row, raw_text, "direct")
                elapsed = time.perf_counter() - started
                result = self.collected_result(index, command_row, raw_file, raw_files, elapsed, attempt + 1, "direct")
                clear_retry_record(log_dir, retry_state, command_row)
                append_event(log_dir, "topas_command_collected", **result)
                return result
            except Exception as exc:
                elapsed = time.perf_counter() - started
                session_error = is_session_error(exc)
                result = {
                    "index": index,
                    "command": command,
                    "commandIdentity": command_identity(command_row),
                    "rawFile": str(raw_file),
                    "rawFiles": [str(path) for path in raw_files],
                    "rawFileCount": len(raw_files),
                    "elapsedSeconds": round(elapsed, 3),
                    "status": "session_error" if session_error else "failed",
                    "attempt": attempt + 1,
                    "error": str(exc),
                    "collectionMode": "direct",
                }
                append_event(log_dir, "topas_session_error" if session_error else "topas_command_failed", **result)
                if session_error:
                    if not continue_on_error:
                        raise TopasSessionError(str(exc)) from exc
                    return result
                if attempt < retries:
                    time.sleep(1)
                    continue
                retry_record_payload = update_retry_failure(
                    log_dir,
                    retry_state,
                    command_row,
                    str(exc),
                    max_collect_error_retries,
                )
                result["retryCount"] = retry_record_payload.get("retryCount")
                result["retryExhausted"] = retry_record_payload.get("retryExhausted")
                if write_error_raw:
                    self.write_error_raw_files(log_dir, command_row, raw_files, exc, retry_record_payload)
                if not continue_on_error:
                    raise
                return result
        raise RuntimeError(f"TOPAS 명령 실행에 실패했습니다: {command}")

    def collect_ac1_group(
        self,
        driver,
        command_rows: list[dict[str, object]],
        log_dir: Path,
        retry_state: dict[str, object],
        start_index: int,
    ) -> list[dict[str, object]]:
        if len(command_rows) <= 1:
            return []
        started = time.perf_counter()
        first_row = command_rows[0]
        ac1_rows = command_rows[1:]
        progress_payload = {
            "startIndex": start_index,
            "commandCount": len(command_rows),
            "ac1Count": len(ac1_rows),
            "expectedCount": len(ac1_rows),
            "firstCommand": str(command_rows[0].get("command") or ""),
            "lastCommand": str(command_rows[-1].get("command") or ""),
        }
        append_event(
            log_dir,
            "topas_ac1_batch_started",
            **progress_payload,
            phase="started",
            seenCount=0,
            inFlightCount=0,
        )
        first_raw = self.execute_command(driver, str(first_row.get("command") or ""))
        self.emit_ac1_progress(log_dir, progress_payload, "first_an_done", seen_count=0, in_flight_count=1)
        ac1_blocks: list[str] = []
        for chunk_start in range(0, len(ac1_rows), TOPAS_AC1_SEND_CHUNK_SIZE):
            chunk_rows = ac1_rows[chunk_start : chunk_start + TOPAS_AC1_SEND_CHUNK_SIZE]
            chunk_payload = dict(progress_payload)
            chunk_payload.update(
                {
                    "chunkStart": chunk_start + 1,
                    "chunkSize": len(chunk_rows),
                    "sentCount": len(ac1_blocks),
                }
            )
            append_event(
                log_dir,
                "topas_ac1_chunk_started",
                **chunk_payload,
            )
            chunk_blocks = self.execute_ac1_batch(
                driver,
                chunk_rows,
                log_dir,
                chunk_payload,
                seen_offset=len(ac1_blocks),
            )
            ac1_blocks.extend(chunk_blocks)
        raw_texts = [first_raw, *ac1_blocks]
        if len(raw_texts) != len(command_rows):
            raise RuntimeError(f"AC1 응답 수 불일치: {len(raw_texts)}/{len(command_rows)}")

        elapsed = time.perf_counter() - started
        self.emit_ac1_progress(
            log_dir,
            progress_payload,
            "raw_writing",
            seen_count=len(ac1_rows),
            in_flight_count=len(command_rows),
        )
        results: list[dict[str, object]] = []
        raw_write_items: list[tuple[dict[str, object], str, str]] = []
        collected_events: list[dict[str, object]] = []
        for offset, (command_row, raw_text) in enumerate(zip(command_rows, raw_texts)):
            raw_files = command_raw_files(command_row)
            raw_file = raw_files[0] if raw_files else Path("")
            collection_mode = "ac1_batch_first" if offset == 0 else "ac1_batch"
            raw_write_items.append((command_row, raw_text, collection_mode))
            result = self.collected_result(
                start_index + offset,
                command_row,
                raw_file,
                raw_files,
                elapsed,
                1,
                collection_mode,
            )
            result["ac1BatchSize"] = len(ac1_rows)
            result["ac1BatchCommandIndex"] = offset
            results.append(result)

        write_started = time.perf_counter()
        self.write_raw_responses_bulk(log_dir, raw_write_items)
        raw_write_seconds = time.perf_counter() - write_started
        self.emit_ac1_progress(
            log_dir,
            progress_payload,
            "raw_written",
            seen_count=len(ac1_rows),
            in_flight_count=len(command_rows),
        )
        clear_retry_records(log_dir, retry_state, command_rows)
        for result in results:
            result["rawWriteSeconds"] = round(raw_write_seconds, 3)
            collected_events.append(make_event("topas_command_collected", **result))
        append_events(log_dir, collected_events)

        append_event(
            log_dir,
            "topas_ac1_batch_collected",
            **progress_payload,
            seenCount=len(ac1_rows),
            inFlightCount=0,
            phase="completed",
            elapsedSeconds=round(elapsed, 3),
            rawWriteSeconds=round(raw_write_seconds, 3),
            rawFileWriteCount=len(raw_write_items),
            rawStorage="sqlite",
        )
        return results

    def collected_result(
        self,
        index: int,
        command_row: dict[str, object],
        raw_file: Path,
        raw_files: list[Path],
        elapsed: float,
        attempt: int,
        collection_mode: str,
    ) -> dict[str, object]:
        return {
            "index": index,
            "command": str(command_row.get("command") or ""),
            "commandIdentity": command_identity(command_row),
            "rawFile": str(raw_file),
            "rawFiles": [str(path) for path in raw_files],
            "rawFileCount": len(raw_files),
            "elapsedSeconds": round(elapsed, 3),
            "status": "collected",
            "attempt": attempt,
            "collectionMode": collection_mode,
        }

    def raw_store_row(
        self,
        command_row: dict[str, object],
        raw_text: str,
        collection_mode: str,
        error: str = "",
    ) -> tuple[object, ...]:
        raw_files = command_raw_files(command_row)
        primary_raw_file = str(raw_files[0]) if raw_files else ""
        return (
            command_identity(command_row),
            primary_raw_file,
            json.dumps([str(path) for path in raw_files], ensure_ascii=False),
            str(command_row.get("routeKey") or ""),
            str(command_row.get("rawKey") or ""),
            str(command_row.get("route") or ""),
            str(command_row.get("direction") or ""),
            str(command_row.get("baseDepartureDate") or ""),
            str(command_row.get("queryDate") or ""),
            str(command_row.get("candidateNights") or ""),
            str(command_row.get("candidateReturnDate") or ""),
            str(command_row.get("command") or ""),
            str(command_row.get("origin") or ""),
            str(command_row.get("destination") or ""),
            str(command_row.get("airline") or ""),
            str(command_row.get("flight") or ""),
            raw_text.strip() + "\n",
            raw_text_state_for_command(command_row, raw_text),
            error,
            collection_mode,
            now_iso(),
        )

    def write_raw_response(
        self,
        log_dir: Path,
        command_row: dict[str, object],
        raw_text: str,
        collection_mode: str,
        error: str = "",
    ) -> None:
        self.write_raw_responses_bulk(log_dir, [(command_row, raw_text, collection_mode, error)])

    def write_raw_responses_bulk(
        self,
        log_dir: Path,
        items: list[tuple[dict[str, object], str, str] | tuple[dict[str, object], str, str, str]],
    ) -> None:
        if not items:
            return
        rows = []
        for item in items:
            command_row, raw_text, collection_mode = item[:3]
            error = item[3] if len(item) > 3 else ""
            rows.append(self.raw_store_row(command_row, raw_text, collection_mode, str(error or "")))
        conn = open_raw_store(log_dir)
        try:
            conn.executemany(
                """
                INSERT OR REPLACE INTO raw_responses (
                    command_identity, primary_raw_file, raw_files_json,
                    route_key, raw_key, route, direction, base_departure_date, query_date,
                    candidate_nights, candidate_return_date, command, origin, destination,
                    airline, flight, raw_text, status, error, collection_mode, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                rows,
            )
            conn.commit()
        finally:
            conn.close()

    def write_raw_files_bulk(self, items: list[tuple[Path, str]]) -> None:
        if not items:
            return
        unique_items: dict[str, tuple[Path, str]] = {}
        for raw_file, raw_text in items:
            if str(raw_file):
                unique_items[str(raw_file)] = (raw_file, raw_text)
        if not unique_items:
            return
        for parent in {raw_file.parent for raw_file, _raw_text in unique_items.values()}:
            parent.mkdir(parents=True, exist_ok=True)

        prepared = [(raw_file, raw_text.strip() + "\n") for raw_file, raw_text in unique_items.values()]
        if len(prepared) == 1:
            raw_file, payload = prepared[0]
            raw_file.write_text(payload, encoding="utf-8")
            return

        max_workers = min(RAW_WRITE_WORKERS, len(prepared))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            list(executor.map(self.write_prepared_raw_file, prepared))

    @staticmethod
    def write_prepared_raw_file(item: tuple[Path, str]) -> None:
        raw_file, payload = item
        raw_file.write_text(payload, encoding="utf-8")

    def write_error_raw_files(
        self,
        log_dir: Path,
        command_row: dict[str, object],
        raw_files: list[Path],
        exc: Exception,
        retry_record_payload: dict[str, object],
    ) -> None:
        command = str(command_row.get("command") or "").strip()
        error_text = (
            "\n".join(
                [
                    f"TOPAS_COLLECT_ERROR {now_iso()}",
                    f"COMMAND: {command}",
                    f"COMMAND_IDENTITY: {command_identity(command_row)}",
                    f"RETRY_COUNT: {retry_record_payload.get('retryCount')}",
                    f"RETRY_EXHAUSTED: {retry_record_payload.get('retryExhausted')}",
                    f"ERROR: {exc}",
                ]
            )
            + "\n"
        )
        self.write_raw_response(log_dir, command_row, error_text, "collect_error", str(exc))

    def emit_ac1_progress(
        self,
        log_dir: Path | None,
        base_payload: dict[str, object] | None,
        phase: str,
        seen_count: int,
        in_flight_count: int,
    ) -> None:
        if log_dir is None:
            return
        payload = dict(base_payload or {})
        payload.update(
            {
                "phase": phase,
                "seenCount": seen_count,
                "inFlightCount": in_flight_count,
            }
        )
        append_event(log_dir, "topas_ac1_batch_progress", **payload)

    def execute_ac1_batch(
        self,
        driver,
        command_rows: list[dict[str, object]],
        log_dir: Path | None = None,
        progress_payload: dict[str, object] | None = None,
        seen_offset: int = 0,
    ) -> list[str]:
        if not command_rows:
            return []
        shell = self.find_element_in_frames(driver, TOPAS_SHELL_ROOT)
        if shell is None:
            raise RuntimeError("TOPAS shell 영역을 찾지 못했습니다.")

        if any(command_query_date(row) is None for row in command_rows):
            raise RuntimeError("AC1 묶음에 날짜가 없는 명령이 포함되어 있습니다.")

        self.emit_ac1_progress(
            log_dir,
            progress_payload,
            "ac1_input_sending",
            seen_count=seen_offset,
            in_flight_count=1 + seen_offset,
        )
        blocks: list[str] = []
        for command_row in command_rows:
            block = self.execute_single_ac1(
                driver,
                command_row,
                log_dir,
                progress_payload,
                seen_offset,
                len(blocks),
                len(command_rows),
            )
            blocks.append(block)
            self.emit_ac1_progress(
                log_dir,
                progress_payload,
                "waiting_response",
                seen_count=seen_offset + len(blocks),
                in_flight_count=1 + seen_offset + len(blocks),
            )

        self.emit_ac1_progress(
            log_dir,
            progress_payload,
            "ac1_input_sent",
            seen_count=seen_offset + len(blocks),
            in_flight_count=1 + seen_offset + len(blocks),
        )
        self.emit_ac1_progress(
            log_dir,
            progress_payload,
            "response_ready",
            seen_count=seen_offset + len(blocks),
            in_flight_count=1 + seen_offset + len(blocks),
        )
        return blocks

    def execute_single_ac1(
        self,
        driver,
        expected_row: dict[str, object],
        log_dir: Path | None,
        progress_payload: dict[str, object] | None,
        seen_offset: int,
        completed_count: int,
        chunk_count: int,
    ) -> str:
        expected_query_date = command_query_date(expected_row)
        if expected_query_date is None:
            raise RuntimeError("AC1 단건 응답에 날짜가 없는 명령이 포함되어 있습니다.")
        expected_date = topas_day_month(expected_query_date)
        expected_command = str(expected_row.get("command") or "")
        last_text = ""
        for retry_count in range(TOPAS_AC1_BLANK_RETRY_LIMIT + 1):
            shell = self.find_element_in_frames(driver, TOPAS_SHELL_ROOT)
            if shell is None:
                raise RuntimeError("TOPAS shell 영역을 찾지 못했습니다.")
            before = self.shell_text(driver, shell)
            before_block_count = len(self.ac1_blocks_for_command(before, expected_row))

            self.send_ac1_keys(driver, 1)
            deadline = time.perf_counter() + min(self.timeout, TOPAS_AC1_SINGLE_RESPONSE_TIMEOUT)
            blank_ready_at: float | None = None
            while time.perf_counter() < deadline:
                shell = self.find_element_in_frames(driver, TOPAS_SHELL_ROOT)
                if shell is None:
                    time.sleep(0.2)
                    continue
                text = self.shell_text(driver, shell)
                last_text = text
                added = self.new_text_after(before, text)
                prompt_ready = self.prompt_is_ready(text)

                added_blocks = self.ac1_blocks_for_command(added, expected_row)
                if added_blocks and prompt_ready:
                    return added_blocks[-1]

                all_blocks = self.ac1_blocks_for_command(text, expected_row)
                if len(all_blocks) > before_block_count and prompt_ready:
                    return all_blocks[-1]

                if prompt_ready and self.has_ac1_echo(added):
                    if blank_ready_at is None:
                        blank_ready_at = time.perf_counter()
                    elif time.perf_counter() - blank_ready_at >= TOPAS_AC1_BLANK_SETTLE_SECONDS:
                        break
                else:
                    blank_ready_at = None
                time.sleep(0.2)

            if retry_count < TOPAS_AC1_BLANK_RETRY_LIMIT:
                if log_dir is not None:
                    payload = dict(progress_payload or {})
                    payload.update(
                        {
                            "expectedDate": expected_date,
                            "seenCount": seen_offset + completed_count,
                            "chunkSeenCount": completed_count,
                            "chunkExpectedCount": chunk_count,
                            "retryCount": retry_count + 1,
                        }
                    )
                    append_event(log_dir, "topas_ac1_blank_retry", **payload)
                continue

        raise TimeoutError(
            f"AC1 단건 응답 대기 초과: {expected_command or expected_date} "
            f"({completed_count}/{chunk_count}개 완료), 마지막 화면 일부: {last_text[-500:]}"
        )

    def ac1_blocks_for_command(self, text: str, expected_row: dict[str, object]) -> list[str]:
        expected = command_signature_from_row(expected_row)
        if expected is None:
            return []
        return [
            block
            for signature, block in self.extract_ac1_response_command_blocks(text)
            if signature == expected
        ]

    def ac1_blocks_for_date(self, text: str, expected_date: str) -> list[str]:
        expected = expected_date.upper()
        return [
            block
            for date_text, block in self.extract_ac1_response_blocks(text)
            if date_text.upper() == expected
        ]

    def has_ac1_echo(self, text: str) -> bool:
        return bool(re.search(r"(?m)^\s*AC1\s*$", text or "", re.IGNORECASE))

    def count_seen_ac1_blocks(self, text: str, expected_dates: list[str]) -> int:
        blocks = self.extract_ac1_response_blocks(text)
        dates = [date_text.upper() for date_text, _block in blocks]
        expected = [date_text.upper() for date_text in expected_dates]
        best = 0
        for start in range(len(dates)):
            matched = 0
            while (
                matched < len(expected)
                and start + matched < len(dates)
                and dates[start + matched] == expected[matched]
            ):
                matched += 1
            best = max(best, matched)
        return best

    def send_ac1_keys(self, driver, count: int) -> None:
        if count <= 0:
            return
        last_error: Exception | None = None
        for index in range(count):
            for attempt in range(TOPAS_AC1_INPUT_RETRIES):
                prompt = self.find_prompt(driver)
                if prompt is None:
                    last_error = RuntimeError("TOPAS 명령 입력창을 찾지 못했습니다.")
                    time.sleep(0.25 + attempt * 0.25)
                    continue
                try:
                    if index == 0 or attempt > 0:
                        self.prepare_prompt(driver, prompt)
                    else:
                        self.focus_prompt(driver, prompt)
                    self.send_single_ac1_key(driver, prompt)
                    time.sleep(TOPAS_AC1_KEY_INTERVAL_SECONDS)
                    break
                except Exception as exc:
                    last_error = exc
                    time.sleep(0.25 + attempt * 0.25)
            else:
                raise RuntimeError(f"AC1 입력 실패: {index + 1}/{count}") from last_error

    def new_text_after(self, before: str, text: str) -> str:
        if text.startswith(before):
            return text[len(before) :]
        for size in (4000, 2000, 1000, 500, 250, 120):
            marker = before[-min(len(before), size) :]
            if not marker:
                continue
            idx = text.rfind(marker)
            if idx >= 0:
                return text[idx + len(marker) :]
        return text

    def extract_ac1_response_command_blocks(self, text: str) -> list[tuple[tuple[str, str, str, str, str], str]]:
        matches = list(TOPAS_AC1_REQUEST_RE.finditer(text))
        blocks: list[tuple[tuple[str, str, str, str, str], str]] = []
        for index, match in enumerate(matches):
            start = match.start()
            end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
            block = text[start:end].strip()
            if any(marker in block.upper() for marker in RAW_COMPLETE_MARKERS):
                blocks.append((command_signature_from_match(match), block))
        return blocks

    def extract_ac1_response_blocks(self, text: str) -> list[tuple[str, str]]:
        return [(signature[0], block) for signature, block in self.extract_ac1_response_command_blocks(text)]

    def match_ac1_blocks(self, text: str, expected_dates: list[str]) -> list[str]:
        blocks = self.extract_ac1_response_blocks(text)
        dates = [date_text for date_text, _block in blocks]
        expected = [date_text.upper() for date_text in expected_dates]
        if len(dates) < len(expected):
            return []
        for start in range(0, len(dates) - len(expected) + 1):
            if dates[start : start + len(expected)] == expected:
                return [block for _date_text, block in blocks[start : start + len(expected)]]
        return []

    def connect(self):
        from selenium import webdriver

        errors: list[str] = []
        for address in self.debugger_addresses:
            driver = None
            try:
                require_topas_debug_target(address)
                options = webdriver.ChromeOptions()
                options.add_experimental_option("debuggerAddress", address)
                driver = webdriver.Chrome(options=options)
                self.window_handle = None
                self.switch_to_topas(driver)
                return driver, address
            except Exception as exc:
                errors.append(f"{address}: {str(exc).splitlines()[0]}")
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:
                        pass
        raise RuntimeError("TOPAS 디버그 브라우저 연결 실패\n" + "\n".join(errors))

    def switch_to_topas(self, driver) -> None:
        handles = list(driver.window_handles)
        if self.window_handle in handles:
            try:
                driver.switch_to.window(self.window_handle)
                driver.switch_to.default_content()
                if self.find_element_in_frames(driver, TOPAS_SHELL_ROOT) is not None:
                    return
            except Exception:
                self.window_handle = None
        matches = []
        for index, handle in enumerate(handles):
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()
                title = (driver.title or "").lower()
                url = (driver.current_url or "").lower()
                score = 0
                if self.find_element_in_frames(driver, TOPAS_SHELL_ROOT) is not None:
                    score += 1000
                if "topassellconnect.com" in url or "topas" in url:
                    score += 100
                if "topas" in title or "sell connect" in title:
                    score += 50
                if score:
                    matches.append((score, -index, handle))
            except Exception:
                continue
        if not matches:
            raise RuntimeError("TOPAS Entry 탭을 찾지 못했습니다.")
        matches.sort(reverse=True)
        self.window_handle = matches[0][2]
        driver.switch_to.window(self.window_handle)
        driver.switch_to.default_content()

    def execute_command(self, driver, command: str) -> str:
        shell = self.find_element_in_frames(driver, TOPAS_SHELL_ROOT)
        if shell is None:
            raise RuntimeError("TOPAS shell 영역을 찾지 못했습니다.")
        before = self.shell_text(driver, shell)

        prompt = self.find_prompt(driver)
        if prompt is None:
            raise RuntimeError("TOPAS 명령 입력창을 찾지 못했습니다.")
        self.prepare_prompt(driver, prompt)
        prompt.send_keys(command)
        prompt.send_keys(self.keys().ENTER)

        deadline = time.perf_counter() + self.timeout
        last_text = ""
        while time.perf_counter() < deadline:
            shell = self.find_element_in_frames(driver, TOPAS_SHELL_ROOT)
            if shell is None:
                time.sleep(0.2)
                continue
            text = self.shell_text(driver, shell)
            last_text = text
            block = self.extract_response_block(text, command, before)
            if block and self.prompt_is_ready(text):
                return block
            time.sleep(0.2)
        raise TimeoutError(f"TOPAS 응답 대기 초과: {command}\n마지막 화면 일부: {last_text[-500:]}")

    def find_prompt(self, driver):
        return self.find_element_in_frames(driver, TOPAS_PROMPT_INPUT, interactable=True)

    def prepare_prompt(self, driver, prompt) -> None:
        focused = driver.execute_script(
            """
            const el = arguments[0];
            el.scrollIntoView({block: 'center', inline: 'center'});
            el.focus();
            el.value = "";
            el.dispatchEvent(new Event('input', {bubbles: true}));
            return document.activeElement === el;
            """,
            prompt,
        )
        if not focused:
            try:
                prompt.click()
            except Exception:
                pass
        time.sleep(0.05)

    def focus_prompt(self, driver, prompt) -> None:
        focused = driver.execute_script(
            """
            const el = arguments[0];
            el.scrollIntoView({block: 'center', inline: 'center'});
            el.focus();
            return document.activeElement === el;
            """,
            prompt,
        )
        if not focused:
            try:
                prompt.click()
            except Exception:
                pass
        time.sleep(0.03)

    def send_single_ac1_key(self, driver, prompt) -> None:
        prompt.send_keys("AC1")
        prompt.send_keys(self.keys().ENTER)

    def find_element_in_frames(self, driver, selector: str, interactable: bool = False, depth: int = 0, max_depth: int = 4):
        try:
            elements = driver.find_elements(self.by().CSS_SELECTOR, selector)
        except Exception:
            elements = []
        for element in elements:
            if not interactable or self.is_interactable(driver, element):
                return element
        if depth >= max_depth:
            return None

        try:
            frames = driver.find_elements(self.by().CSS_SELECTOR, "iframe, frame")
        except Exception:
            frames = []
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                found = self.find_element_in_frames(driver, selector, interactable, depth + 1, max_depth)
                if found is not None:
                    return found
            except Exception:
                pass
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass
        return None

    def is_interactable(self, driver, element) -> bool:
        try:
            if not element.is_displayed() or not element.is_enabled():
                return False
            return bool(
                driver.execute_script(
                    """
                    const el = arguments[0];
                    if (!el || !el.isConnected) return false;
                    const rect = el.getBoundingClientRect();
                    const style = window.getComputedStyle(el);
                    return rect.width > 0 &&
                        rect.height > 0 &&
                        style.display !== 'none' &&
                        style.visibility !== 'hidden' &&
                        style.pointerEvents !== 'none' &&
                        !el.disabled &&
                        !el.readOnly;
                    """,
                    element,
                )
            )
        except Exception:
            return False

    def shell_text(self, driver, shell) -> str:
        try:
            text = driver.execute_script("return arguments[0].innerText || arguments[0].textContent || '';", shell) or ""
            if text.strip():
                return text
        except Exception:
            pass
        return shell.text or ""

    def extract_response_block(self, text: str, command: str, before: str) -> str:
        idx = self.find_response_start(text, command, before)
        if idx < 0:
            return ""
        block = text[idx:].strip()
        upper = block.upper()
        if "AMADEUS AVAILABILITY" not in upper and "NO FLIGHT" not in upper and "REQUEST NEW AVAILABILITY" not in upper:
            return ""
        search_start = len(command) if block.upper().startswith(command.upper()) else 0
        next_command = re.search(
            r"\n\s*AN\d{1,2}[A-Z]{3}[A-Z]{6}/A[A-Z0-9]{2}\d{0,4}",
            block[search_start:],
            re.IGNORECASE,
        )
        if next_command:
            block = block[: search_start + next_command.start()]
        return block.strip()

    def find_response_start(self, text: str, command: str, before: str) -> int:
        idx = text.rfind(command)
        if idx >= 0:
            return idx

        expected_signature = command_signature_from_text(command)
        if expected_signature is not None:
            matches = [
                match
                for match in TOPAS_FULL_COMMAND_RE.finditer(text)
                if command_signature_from_match(match) == expected_signature
            ]
            if matches:
                return matches[-1].start()

        command_match = TOPAS_COMMAND_RE.search(command)
        if command_match:
            command_date = command_match.group("date").upper()
            command_destination = command_match.group("destination").upper()
            matches = [
                match
                for match in TOPAS_AVAILABILITY_HEADER_RE.finditer(text)
                if match.group("date").upper() == command_date
                and match.group("destination").upper() == command_destination
            ]
            if matches:
                return matches[-1].start()

        if before and text.startswith(before):
            idx = len(before)
            if idx < len(text):
                return idx

        marker_indexes = [
            text.upper().rfind("** AMADEUS AVAILABILITY"),
            text.upper().rfind("REQUEST NEW AVAILABILITY"),
            text.upper().rfind("NO FLIGHT"),
        ]
        idx = max(marker_indexes)
        if idx >= 0:
            line_start = text.rfind("\n", 0, idx)
            return 0 if line_start < 0 else line_start + 1
        return -1

    def prompt_is_ready(self, text: str) -> bool:
        tail = "\n".join(str(text or "").replace("\r\n", "\n").splitlines()[-8:])
        return bool(PROMPT_READY_RE.search(tail)) and not LOADING_RE.search(tail)

    @staticmethod
    def by():
        from selenium.webdriver.common.by import By

        return By

    @staticmethod
    def keys():
        from selenium.webdriver.common.keys import Keys

        return Keys


def retry_final_pending_commands(log_dir: Path, args: argparse.Namespace, collector: TopasLiveCollector) -> list[dict[str, object]]:
    pending_before = collect_pending_summary(log_dir, args.max_collect_error_retries)
    commands = select_commands(
        log_dir,
        args.direction,
        "raw-missing",
        True,
        0,
        excluded_ids=None,
        retry_collect_errors=True,
        max_collect_error_retries=args.max_collect_error_retries,
    )
    append_event(
        log_dir,
        "topas_final_pending_retry_started",
        pendingSummary=pending_before,
        commandCount=len(commands),
    )
    if not commands:
        append_event(
            log_dir,
            "topas_final_pending_retry_finished",
            commandCount=0,
            collectedCount=0,
            failedCount=0,
            sessionErrorCount=0,
            pendingSummary=collect_pending_summary(log_dir, args.max_collect_error_retries),
        )
        return []

    print(f"최종 엑셀 생성 전 남은 TOPAS 조회 {len(commands)}건을 한 번 더 재시도합니다.")
    try:
        results = collector.collect(
            commands,
            log_dir,
            continue_on_error=True,
            retries=args.retries,
            write_error_raw=True,
            max_collect_error_retries=args.max_collect_error_retries,
            session_error_threshold=args.session_error_threshold,
        )
    except Exception as exc:
        append_event(log_dir, "topas_final_pending_retry_failed", error=str(exc))
        print(f"최종 pending 재시도 중 오류가 발생했습니다: {exc}")
        return []

    pending_after = collect_pending_summary(log_dir, args.max_collect_error_retries)
    append_event(
        log_dir,
        "topas_final_pending_retry_finished",
        commandCount=len(commands),
        collectedCount=sum(1 for row in results if row.get("status") == "collected"),
        failedCount=sum(1 for row in results if row.get("status") == "failed"),
        sessionErrorCount=sum(1 for row in results if row.get("status") == "session_error"),
        pendingSummary=pending_after,
    )
    return results


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="항공자동조회 TOPAS live collector")
    parser.add_argument("log_dir", help="output/logs/{runId} 폴더")
    parser.add_argument("--direction", choices=["departure", "return", "all"], default="return")
    parser.add_argument("--limit", type=int, default=DEFAULT_COLLECT_BATCH_SIZE, help="내부 체크포인트 배치 크기")
    parser.add_argument("--debugger-address", action="append", default=[])
    parser.add_argument("--timeout", type=float, default=80.0)
    parser.add_argument("--live", action="store_true", help="실제 TOPAS 디버그 크롬에 명령을 전송")
    parser.add_argument(
        "--filter",
        dest="command_filter",
        choices=["all", "raw-missing", "unconfirmed", "fare-route-missing"],
        default="all",
        help="실행할 TOPAS 명령을 상태별로 추림",
    )
    parser.add_argument("--skip-existing", action="store_true", help="이미 raw 파일이 있으면 건너뜀")
    parser.add_argument("--process-after", action="store_true", help="수집 후 raw 결과 JSON 후처리")
    parser.add_argument("--calculate-after", action="store_true", help="수집 후 운임 DB 계산까지 실행")
    parser.add_argument("--loop-until-done", action="store_true", help="남은 명령이 없을 때까지 반복. limit는 내부 체크포인트 크기로만 사용")
    parser.add_argument("--batch-pause", type=float, default=1.0, help="연속 실행 배치 사이 대기 초")
    parser.add_argument("--max-batches", type=int, default=0, help="테스트용: 연속 실행 최대 배치 수, 0이면 제한 없음")
    parser.add_argument("--retries", type=int, default=0, help="개별 TOPAS 명령 실패 시 재시도 횟수")
    parser.add_argument("--continue-on-error", action="store_true", help="개별 명령 실패 시 다음 명령으로 계속 진행")
    parser.add_argument("--write-error-raw", action="store_true", help="개별 명령 최종 실패 시 error raw를 써서 다음 루프에서 건너뜀")
    parser.add_argument("--retry-collect-errors", action="store_true", help="TOPAS_COLLECT_ERROR/불완전 raw를 retry-state 기준으로 재수집")
    parser.add_argument("--max-collect-error-retries", type=int, default=DEFAULT_MAX_COLLECT_ERROR_RETRIES)
    parser.add_argument("--session-error-threshold", type=int, default=DEFAULT_SESSION_ERROR_THRESHOLD)
    args = parser.parse_args(argv)

    log_dir = Path(args.log_dir)
    if args.retries < 0:
        raise SystemExit("--retries는 0 이상이어야 합니다.")
    if args.max_collect_error_retries < 1:
        raise SystemExit("--max-collect-error-retries는 1 이상이어야 합니다.")
    if args.session_error_threshold < 1:
        raise SystemExit("--session-error-threshold는 1 이상이어야 합니다.")

    addresses = args.debugger_address or ["127.0.0.1:9222"]
    effective_skip_existing = args.skip_existing or args.loop_until_done
    processed_ids: set[str] = set()
    all_results: list[dict[str, object]] = []
    batch_index = 0
    no_remaining_commands = False
    stopped_by_session_error = False
    stopped_by_user = False
    stopped_by_error = False
    consecutive_session_error_batches = 0
    resumed_after_session_error = read_run_status(log_dir) == "session_error"
    collector = TopasLiveCollector(addresses, timeout=args.timeout)
    run_after_each_batch = (args.process_after or args.calculate_after) and not args.loop_until_done
    while True:
        if wait_if_paused(log_dir):
            stopped_by_user = True
            break
        if stop_requested(log_dir):
            stopped_by_user = True
            append_event(log_dir, "topas_collection_stop_requested", phase="before_batch")
            print("정지 요청 확인: 새 TOPAS 묶음을 시작하지 않고 종료합니다.")
            break

        commands = select_commands(
            log_dir,
            args.direction,
            args.command_filter,
            effective_skip_existing,
            args.limit,
            processed_ids,
            retry_collect_errors=args.retry_collect_errors,
            max_collect_error_retries=args.max_collect_error_retries,
        )
        if not commands:
            if batch_index == 0:
                raise SystemExit("실행할 TOPAS 명령이 없습니다.")
            print("남은 TOPAS 명령이 없습니다.")
            append_event(log_dir, "topas_collection_no_remaining", totalBatches=batch_index)
            no_remaining_commands = True
            break

        batch_index += 1
        if not args.live:
            print(f"[dry-run batch {batch_index}] TOPAS 명령 {len(commands)}개")
            for row in commands:
                raw_files = command_raw_files(row)
                suffix = f" (raw {len(raw_files)}개)" if len(raw_files) > 1 else ""
                states = ",".join(raw_file_state(path) for path in raw_files)
                print(f"{row.get('command')} -> {row.get('rawFile')}{suffix} [{states}]")
            processed_ids.update(command_identity(row) for row in commands)
            if not args.loop_until_done:
                break
            if args.max_batches > 0 and batch_index >= args.max_batches:
                print(f"--max-batches {args.max_batches} 도달로 중지합니다.")
                break
            continue

        append_event(
            log_dir,
            "topas_batch_started",
            batchIndex=batch_index,
            direction=args.direction,
            commandFilter=args.command_filter,
            batchSize=len(commands),
            loopUntilDone=args.loop_until_done,
            retryCollectErrors=args.retry_collect_errors,
            maxCollectErrorRetries=args.max_collect_error_retries,
        )
        print(f"[batch {batch_index}] TOPAS 명령 {len(commands)}개 실행")
        try:
            results = collector.collect(
                commands,
                log_dir,
                continue_on_error=args.continue_on_error or args.loop_until_done,
                retries=args.retries,
                write_error_raw=args.write_error_raw or args.loop_until_done,
                max_collect_error_retries=args.max_collect_error_retries,
                session_error_threshold=args.session_error_threshold,
            )
        except TopasSessionError as exc:
            stopped_by_session_error = True
            results = []
            append_event(log_dir, "topas_collection_stopped", reason="session_error", error=str(exc))
            print(f"TOPAS 세션 오류로 수집을 중단합니다: {exc}")
            if run_after_each_batch:
                run_after_steps(log_dir, args.process_after, args.calculate_after, write_excel=False)
            break
        except Exception as exc:
            stopped_by_session_error = is_session_error(exc)
            results = []
            status = "session_error" if stopped_by_session_error else "interrupted"
            if not stopped_by_session_error:
                stopped_by_error = True
            append_event(log_dir, "topas_collection_stopped", reason="unexpected_error", error=str(exc))
            update_run_status(log_dir, status, lastError=str(exc))
            print(f"TOPAS 수집 중 예외가 발생해 중단합니다: {exc}")
            if run_after_each_batch:
                run_after_steps(log_dir, args.process_after, args.calculate_after, write_excel=False)
            break
        all_results.extend(results)
        processed_ids.update(command_identity(row) for row in commands)
        session_error_count = sum(1 for row in results if row.get("status") == "session_error")
        if session_error_count and session_error_count == len(results):
            consecutive_session_error_batches += session_error_count
        else:
            consecutive_session_error_batches = 0
        if any(row.get("status") == "collected" for row in results):
            status_updates = {"lastError": ""}
            if resumed_after_session_error:
                status_updates["resumedAfterSessionError"] = True
                resumed_after_session_error = False
            update_run_status(log_dir, "running", **status_updates)
        append_event(
            log_dir,
            "topas_batch_finished",
            batchIndex=batch_index,
            collectedCount=sum(1 for row in results if row.get("status") == "collected"),
            failedCount=sum(1 for row in results if row.get("status") == "failed"),
            sessionErrorCount=session_error_count,
        )
        print(json.dumps(results, ensure_ascii=False, indent=2))

        if run_after_each_batch:
            run_after_steps(log_dir, args.process_after, args.calculate_after, write_excel=False)

        if consecutive_session_error_batches >= args.session_error_threshold:
            message = f"TOPAS 세션 오류가 {consecutive_session_error_batches}회 연속 발생해 수집을 중단합니다."
            append_event(
                log_dir,
                "topas_session_error_circuit_opened",
                threshold=args.session_error_threshold,
                error=message,
            )
            update_run_status(log_dir, "session_error", lastError=message)
            stopped_by_session_error = True
            print(message)
            break

        if stop_requested(log_dir):
            stopped_by_user = True
            append_event(log_dir, "topas_collection_stopped", reason="user_stop", batchIndex=batch_index)
            print("정지 요청 확인: 현재 묶음 저장 후 종료합니다.")
            break

        if not args.loop_until_done:
            break
        if args.max_batches > 0 and batch_index >= args.max_batches:
            print(f"--max-batches {args.max_batches} 도달로 중지합니다.")
            break
        if args.batch_pause > 0:
            time.sleep(args.batch_pause)

    if not args.live:
        print("--live 옵션을 주면 실제 TOPAS에 전송합니다.")

    pending = collect_pending_summary(log_dir, args.max_collect_error_retries)
    if stopped_by_user:
        update_run_status(log_dir, "interrupted", controlState="stopped", stoppedByUser=True, pendingSummary=pending)
    elif args.live and no_remaining_commands and not stopped_by_session_error and (args.process_after or args.calculate_after):
        if actionable_pending_count(pending) > 0:
            all_results.extend(retry_final_pending_commands(log_dir, args, collector))
            pending = collect_pending_summary(log_dir, args.max_collect_error_retries)

        if actionable_pending_count(pending) == 0:
            print("자동 재시도 가능한 수집 대상이 없습니다. 최종 후처리/운임 계산을 실행합니다.")
            append_event(
                log_dir,
                "topas_final_processing_started",
                processAfter=args.process_after or args.calculate_after,
                calculateAfter=args.calculate_after,
                writeExcel=args.calculate_after,
                forcedExcelWithPending=False,
            )
            run_after_steps(
                log_dir,
                process_after=args.process_after or args.calculate_after,
                calculate_after=args.calculate_after,
                write_excel=args.calculate_after,
                force_excel_with_pending=False,
            )
            append_event(
                log_dir,
                "topas_final_processing_finished",
                processAfter=args.process_after or args.calculate_after,
                calculateAfter=args.calculate_after,
                writeExcel=args.calculate_after,
                forcedExcelWithPending=False,
            )
        elif args.calculate_after:
            print(
                f"재시도 후에도 actionable pending {pending.get('actionablePendingCommands')}건이 남아 "
                "조회오류/미수집 표시로 최종 엑셀을 생성합니다."
            )
            append_event(
                log_dir,
                "topas_final_processing_started",
                processAfter=True,
                calculateAfter=True,
                writeExcel=True,
                forcedExcelWithPending=True,
                pendingSummary=pending,
            )
            run_after_steps(
                log_dir,
                process_after=True,
                calculate_after=True,
                write_excel=True,
                force_excel_with_pending=True,
            )
            append_event(
                log_dir,
                "topas_final_processing_finished",
                processAfter=True,
                calculateAfter=True,
                writeExcel=True,
                forcedExcelWithPending=True,
                pendingSummary=collect_pending_summary(log_dir, args.max_collect_error_retries),
            )
        else:
            print(f"최종 엑셀은 아직 생성하지 않습니다. actionable pending {pending.get('actionablePendingCommands')}건")
            update_run_status(log_dir, "interrupted", excelStatus="pending_final_collection", pendingSummary=pending)
    elif args.live and not stopped_by_session_error and actionable_pending_count(pending) > 0:
        update_run_status(log_dir, "interrupted", excelStatus="pending_final_collection", pendingSummary=pending)

    collector.close()

    print(
        json.dumps(
            {
                "totalBatches": batch_index,
                "totalCommands": len(all_results) if args.live else len(processed_ids),
                "collectedCount": sum(1 for row in all_results if row.get("status") == "collected"),
                "failedCount": sum(1 for row in all_results if row.get("status") == "failed"),
                "sessionErrorCount": sum(1 for row in all_results if row.get("status") == "session_error"),
                "pendingSummary": collect_pending_summary(log_dir, args.max_collect_error_retries),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    if stopped_by_session_error:
        return 3
    if stopped_by_error:
        return 4
    return 0


def select_commands(
    log_dir: Path,
    direction: str,
    command_filter: str,
    skip_existing: bool,
    limit: int,
    excluded_ids: set[str] | None = None,
    retry_collect_errors: bool = False,
    max_collect_error_retries: int = DEFAULT_MAX_COLLECT_ERROR_RETRIES,
) -> list[dict[str, object]]:
    commands = load_plan(log_dir, direction)
    commands = merge_duplicate_commands(commands)
    retry_state = load_retry_state(log_dir)
    raw_records = load_raw_store_records(log_dir)
    commands = filter_commands(
        log_dir,
        commands,
        command_filter,
        retry_state,
        retry_collect_errors,
        max_collect_error_retries,
        raw_records,
    )
    if excluded_ids:
        commands = [row for row in commands if command_identity(row) not in excluded_ids]
    if skip_existing:
        commands = [
            row
            for row in commands
            if command_has_actionable_raw(
                log_dir,
                row,
                retry_state,
                retry_collect_errors,
                max_collect_error_retries,
                raw_records,
            )
        ]
    if limit > 0:
        commands = commands[:limit]
    return commands


def count_missing_raw_files(log_dir: Path) -> int:
    return int(collect_pending_summary(log_dir).get("actionablePendingCommands") or 0)


def run_after_steps(
    log_dir: Path,
    process_after: bool,
    calculate_after: bool,
    write_excel: bool = False,
    force_excel_with_pending: bool = False,
) -> None:
    from air_auto_lookup_mvp import calculate_run_fares, process_run_raw_outputs

    outputs = process_run_raw_outputs(log_dir)
    if process_after:
        for path in outputs.values():
            print(path)
    if calculate_after:
        fare_results = calculate_run_fares(
            log_dir,
            write_excel=write_excel,
            force_excel_with_pending=force_excel_with_pending,
        )
        print(fare_results)


if __name__ == "__main__":
    raise SystemExit(main())
