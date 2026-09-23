from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import StringIO
import json
import os
from pathlib import Path
import sqlite3
from uuid import uuid4

import pandas as pd


@dataclass(frozen=True)
class HistoryRecordSummary:
    record_id: str
    created_at: str
    reconciliation_type: str
    title: str
    input_files: tuple[str, ...]
    status_counts: dict[str, int]
    report_filename: str


@dataclass(frozen=True)
class HistoryRecord(HistoryRecordSummary):
    summary: dict[str, object]
    results: pd.DataFrame
    issues: pd.DataFrame
    report_bytes: bytes


def default_history_database() -> Path:
    configured_path = os.environ.get("DOI_CHIEU_HISTORY_DB", "").strip()
    if configured_path:
        return Path(configured_path)
    return Path(__file__).resolve().parents[1] / "data" / "tool_history.sqlite3"


def _connect(database: str | Path) -> sqlite3.Connection:
    path = Path(database)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA busy_timeout = 10000")
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS reconciliation_history (
            record_id TEXT PRIMARY KEY,
            created_at TEXT NOT NULL,
            reconciliation_type TEXT NOT NULL,
            title TEXT NOT NULL,
            input_files_json TEXT NOT NULL,
            status_counts_json TEXT NOT NULL,
            summary_json TEXT NOT NULL,
            results_json TEXT NOT NULL,
            issues_json TEXT NOT NULL,
            report_filename TEXT NOT NULL,
            report_blob BLOB NOT NULL
        )
        """
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_history_created_at "
        "ON reconciliation_history(created_at DESC)"
    )
    connection.commit()
    return connection


def _dataframe_to_json(dataframe: pd.DataFrame) -> str:
    return dataframe.to_json(orient="split", force_ascii=False)


def _dataframe_from_json(payload: str) -> pd.DataFrame:
    if not payload:
        return pd.DataFrame()
    return pd.read_json(StringIO(payload), orient="split")


def save_history_record(
    *,
    reconciliation_type: str,
    title: str,
    input_files: list[str] | tuple[str, ...],
    status_counts: dict[str, int],
    summary: dict[str, object],
    results: pd.DataFrame,
    issues: pd.DataFrame,
    report_filename: str,
    report_bytes: bytes,
    database: str | Path | None = None,
) -> str:
    record_id = uuid4().hex
    created_at = datetime.now().astimezone().isoformat(timespec="seconds")
    database_path = database or default_history_database()
    with _connect(database_path) as connection:
        connection.execute(
            """
            INSERT INTO reconciliation_history (
                record_id, created_at, reconciliation_type, title,
                input_files_json, status_counts_json, summary_json,
                results_json, issues_json, report_filename, report_blob
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                record_id,
                created_at,
                reconciliation_type,
                title,
                json.dumps(list(input_files), ensure_ascii=False),
                json.dumps(status_counts, ensure_ascii=False),
                json.dumps(summary, ensure_ascii=False, default=str),
                _dataframe_to_json(results),
                _dataframe_to_json(issues),
                report_filename,
                sqlite3.Binary(report_bytes),
            ),
        )
    return record_id


def list_history_records(
    database: str | Path | None = None,
    *,
    limit: int = 100,
) -> list[HistoryRecordSummary]:
    database_path = database or default_history_database()
    with _connect(database_path) as connection:
        rows = connection.execute(
            """
            SELECT record_id, created_at, reconciliation_type, title,
                   input_files_json, status_counts_json, report_filename
            FROM reconciliation_history
            ORDER BY created_at DESC, rowid DESC
            LIMIT ?
            """,
            (max(1, int(limit)),),
        ).fetchall()
    return [
        HistoryRecordSummary(
            record_id=row["record_id"],
            created_at=row["created_at"],
            reconciliation_type=row["reconciliation_type"],
            title=row["title"],
            input_files=tuple(json.loads(row["input_files_json"])),
            status_counts={
                str(key): int(value)
                for key, value in json.loads(row["status_counts_json"]).items()
            },
            report_filename=row["report_filename"],
        )
        for row in rows
    ]


def get_history_record(
    record_id: str,
    database: str | Path | None = None,
) -> HistoryRecord | None:
    database_path = database or default_history_database()
    with _connect(database_path) as connection:
        row = connection.execute(
            "SELECT * FROM reconciliation_history WHERE record_id = ?",
            (record_id,),
        ).fetchone()
    if row is None:
        return None
    return HistoryRecord(
        record_id=row["record_id"],
        created_at=row["created_at"],
        reconciliation_type=row["reconciliation_type"],
        title=row["title"],
        input_files=tuple(json.loads(row["input_files_json"])),
        status_counts={
            str(key): int(value)
            for key, value in json.loads(row["status_counts_json"]).items()
        },
        report_filename=row["report_filename"],
        summary=json.loads(row["summary_json"]),
        results=_dataframe_from_json(row["results_json"]),
        issues=_dataframe_from_json(row["issues_json"]),
        report_bytes=bytes(row["report_blob"]),
    )


def delete_history_record(
    record_id: str,
    database: str | Path | None = None,
) -> bool:
    """Delete one saved reconciliation and its report.

    Returns ``True`` only when the requested record existed. The parameterized
    query keeps the operation scoped to the exact immutable record id.
    """
    database_path = database or default_history_database()
    with _connect(database_path) as connection:
        cursor = connection.execute(
            "DELETE FROM reconciliation_history WHERE record_id = ?",
            (record_id,),
        )
    return cursor.rowcount == 1


def delete_all_history_records(database: str | Path | None = None) -> int:
    """Delete every saved reconciliation and return the deleted row count."""
    database_path = database or default_history_database()
    with _connect(database_path) as connection:
        cursor = connection.execute("DELETE FROM reconciliation_history")
    return max(0, cursor.rowcount)
