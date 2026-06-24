from __future__ import annotations

import pickle
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any

import gspread

from process_sprinklr import (
    DEFAULT_DATA_SHEET,
    audit_entry,
    is_deleted_post,
    normalize_permalink,
    read_sprinklr_rows,
    transform_sprinklr_row,
)


BASE_DIR = Path(__file__).resolve().parent
TOKEN_PATH = BASE_DIR / "token.pickle"


def authorize_client(token_path: Path = TOKEN_PATH):
    with token_path.open("rb") as token:
        creds = pickle.load(token)
    return gspread.authorize(creds)


def open_worksheet(sheet_url: str, worksheet_name: str, token_path: Path = TOKEN_PATH):
    client = authorize_client(token_path)
    sheet = client.open_by_url(sheet_url)
    return sheet.worksheet(worksheet_name)


def normalize_google_header(value: Any) -> str:
    return "" if value is None else str(value).strip()


def get_google_sheet_headers(worksheet) -> list[str]:
    return [normalize_google_header(value) for value in worksheet.row_values(1)]


def google_header_index(headers: list[str]) -> dict[str, int]:
    return {header: idx + 1 for idx, header in enumerate(headers) if header}


def get_google_permalink_index(worksheet, headers: list[str] | None = None) -> dict[str, int]:
    headers = headers or get_google_sheet_headers(worksheet)
    header_index = google_header_index(headers)
    if "Permalink" not in header_index:
        raise ValueError("Google Sheet is missing required header: Permalink")

    permalink_col = header_index["Permalink"]
    index: dict[str, int] = {}
    for row_number, value in enumerate(worksheet.col_values(permalink_col)[1:], start=2):
        permalink = normalize_permalink(value)
        if permalink:
            index[permalink] = row_number
    return index


def _trim_number_text(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.6f}".rstrip("0").rstrip(".")


def _format_percent_value(value: Any) -> Any:
    if value in (None, ""):
        return ""
    if isinstance(value, str):
        text = value.strip()
        if text.upper() == "NA":
            return "NA"
        if text.endswith("%"):
            return text
        try:
            return f"{_trim_number_text(float(text))}%"
        except ValueError:
            return value
    try:
        number = float(value)
    except (TypeError, ValueError):
        return value

    # process_sprinklr converts some Sprinklr percent strings into decimals.
    # Convert those back to readable percentage text for Google Sheets.
    if abs(number) <= 1:
        number = number * 100
    return f"{_trim_number_text(number)}%"


def _format_seconds_value(value: Any) -> Any:
    if value in (None, ""):
        return ""
    if isinstance(value, str):
        text = value.strip()
        if text.upper() == "NA":
            return "NA"
        if ":" in text:
            parts = text.split(":")
            try:
                if len(parts) == 3:
                    hours = float(parts[0])
                    minutes = float(parts[1])
                    seconds = float(parts[2])
                    return _trim_number_text(hours * 3600 + minutes * 60 + seconds)
            except ValueError:
                return value
        try:
            return _trim_number_text(float(text))
        except ValueError:
            return value
    if isinstance(value, timedelta):
        return _trim_number_text(value.total_seconds())
    if isinstance(value, time):
        seconds = value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
        return _trim_number_text(seconds)
    try:
        return _trim_number_text(float(value))
    except (TypeError, ValueError):
        return value


def _format_number_value(value: Any) -> Any:
    if value in (None, ""):
        return ""
    if isinstance(value, str):
        text = value.strip()
        if text.upper() == "NA":
            return "NA"
        try:
            number = float(text.replace(",", ""))
            return int(number) if number.is_integer() else number
        except ValueError:
            return value
    if isinstance(value, timedelta):
        return value.total_seconds()
    return value


def format_for_google_sheet(value: Any, header: str) -> Any:
    if value in (None, ""):
        return ""
    if header in {"Skip Rate %", "Share rate", "ER%"}:
        return _format_percent_value(value)
    if header == "Avg Watch Time (Seconds)":
        return _format_seconds_value(value)
    if isinstance(value, datetime):
        if header == "Date":
            return value.strftime("%d-%m-%Y")
        if header == "Time":
            return value.strftime("%H:%M")
        return value.strftime("%d-%m-%Y %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if header in {
        "Reach",
        "Views",
        "PE",
        "Likes",
        "Saves",
        "Shares",
        "Reposts",
        "Comments",
        "Pos Comments",
        "Neg Comments",
        "Neutral Comments",
        "Vis Engagement",
        "Follows",
    }:
        return _format_number_value(value)
    return _format_number_value(value)


def transformed_row_to_google_values(transformed_row: dict[str, Any], headers: list[str]) -> list[Any]:
    return [
        format_for_google_sheet(transformed_row.get(header, ""), header)
        for header in headers
    ]


def transformed_metric_updates(transformed_row: dict[str, Any], headers: list[str], update_columns: list[str]) -> dict[str, Any]:
    available_headers = set(headers)
    return {
        header: format_for_google_sheet(transformed_row.get(header, ""), header)
        for header in update_columns
        if header in available_headers and header in transformed_row
    }


def validate_google_headers(headers: list[str], config: dict[str, Any], mode: str) -> list[str]:
    required = {"Permalink"}
    if mode == "weekly":
        required.update(config["sprinklr_to_master"].values())
    else:
        required.update(config.get("periodical_update_columns", []))
    available = set(headers)
    return sorted(header for header in required if header not in available)


def _read_transformable_rows(sprinklr_path: Path, config: dict[str, Any]) -> list[dict[str, Any]]:
    return read_sprinklr_rows(
        sprinklr_path,
        int(config.get("sprinklr_header_row", 3)),
    )


def weekly_direct_append(
    sprinklr_path: str | Path,
    sheet_url: str,
    worksheet_name: str,
    config: dict[str, Any],
    dry_run: bool = True,
) -> dict[str, Any]:
    sprinklr_path = Path(sprinklr_path)
    worksheet = open_worksheet(sheet_url, worksheet_name)
    headers = get_google_sheet_headers(worksheet)
    missing_headers = validate_google_headers(headers, config, "weekly")
    if missing_headers:
        raise ValueError(f"Google Sheet is missing required headers: {missing_headers}")

    existing = get_google_permalink_index(worksheet, headers)
    rows = _read_transformable_rows(sprinklr_path, config)
    rows_to_append: list[list[Any]] = []
    row_audit: list[dict[str, Any]] = []

    deleted_posts_removed = 0
    missing_permalink_rows = 0
    duplicate_rows_skipped = 0

    for sprinklr_row in rows:
        source_row = int(sprinklr_row.get("_source_row", 0))
        permalink = normalize_permalink(sprinklr_row.get("Permalink"))

        if is_deleted_post(sprinklr_row, config.get("deleted_post_token", "deleted post")):
            deleted_posts_removed += 1
            row_audit.append(audit_entry(source_row, "removed", permalink, "Deleted post"))
            continue
        if not permalink:
            missing_permalink_rows += 1
            row_audit.append(audit_entry(source_row, "skipped", permalink, "Missing permalink"))
            continue
        if permalink in existing:
            duplicate_rows_skipped += 1
            row_audit.append(audit_entry(source_row, "skipped", permalink, "Permalink already exists", existing[permalink]))
            continue

        transformed = transform_sprinklr_row(sprinklr_row, config)
        rows_to_append.append(transformed_row_to_google_values(transformed, headers))
        row_audit.append(audit_entry(source_row, "appended", permalink, "New permalink"))

    start_row = None
    if rows_to_append and not dry_run:
        permalink_col = google_header_index(headers)["Permalink"]
        start_row = len(worksheet.col_values(permalink_col)) + 1
        worksheet.update(
            values=rows_to_append,
            range_name=f"A{start_row}",
            value_input_option="USER_ENTERED",
        )

    return {
        "mode": "weekly_direct_append",
        "dry_run": dry_run,
        "sprinklr_path": str(sprinklr_path),
        "worksheet_name": worksheet_name,
        "google_rows_found": len(existing),
        "sprinklr_rows_seen": len(rows),
        "deleted_posts_removed": deleted_posts_removed,
        "missing_permalink_rows": missing_permalink_rows,
        "duplicate_rows_skipped": duplicate_rows_skipped,
        "rows_to_append": len(rows_to_append),
        "rows_appended": 0 if dry_run else len(rows_to_append),
        "start_row": start_row,
        "missing_google_headers": [],
        "row_audit": row_audit,
    }


def periodical_direct_update(
    sprinklr_path: str | Path,
    sheet_url: str,
    worksheet_name: str,
    config: dict[str, Any],
    dry_run: bool = True,
) -> dict[str, Any]:
    sprinklr_path = Path(sprinklr_path)
    worksheet = open_worksheet(sheet_url, worksheet_name)
    headers = get_google_sheet_headers(worksheet)
    missing_headers = validate_google_headers(headers, config, "periodical")
    if missing_headers:
        raise ValueError(f"Google Sheet is missing required headers: {missing_headers}")

    header_index = google_header_index(headers)
    existing = get_google_permalink_index(worksheet, headers)
    rows = _read_transformable_rows(sprinklr_path, config)
    row_audit: list[dict[str, Any]] = []
    updates: list[dict[str, Any]] = []

    deleted_posts_removed = 0
    missing_permalink_rows = 0
    unmatched_periodical_rows = 0

    for sprinklr_row in rows:
        source_row = int(sprinklr_row.get("_source_row", 0))
        permalink = normalize_permalink(sprinklr_row.get("Permalink"))

        if is_deleted_post(sprinklr_row, config.get("deleted_post_token", "deleted post")):
            deleted_posts_removed += 1
            row_audit.append(audit_entry(source_row, "removed", permalink, "Deleted post"))
            continue
        if not permalink:
            missing_permalink_rows += 1
            row_audit.append(audit_entry(source_row, "skipped", permalink, "Missing permalink"))
            continue

        target_row = existing.get(permalink)
        if target_row is None:
            unmatched_periodical_rows += 1
            row_audit.append(audit_entry(source_row, "unmatched", permalink, "Permalink not found"))
            continue

        transformed = transform_sprinklr_row(sprinklr_row, config)
        metric_values = transformed_metric_updates(
            transformed,
            headers,
            config.get("periodical_update_columns", []),
        )
        for header, value in metric_values.items():
            updates.append(
                {
                    "range": f"{gspread.utils.rowcol_to_a1(target_row, header_index[header])}",
                    "values": [[value]],
                }
            )
        row_audit.append(audit_entry(source_row, "updated", permalink, "Matched permalink", target_row))

    if updates and not dry_run:
        worksheet.batch_update(updates, value_input_option="USER_ENTERED")

    return {
        "mode": "periodical_direct_update",
        "dry_run": dry_run,
        "sprinklr_path": str(sprinklr_path),
        "worksheet_name": worksheet_name,
        "google_rows_found": len(existing),
        "sprinklr_rows_seen": len(rows),
        "deleted_posts_removed": deleted_posts_removed,
        "missing_permalink_rows": missing_permalink_rows,
        "unmatched_periodical_rows": unmatched_periodical_rows,
        "rows_to_update": len({entry["target_row"] for entry in row_audit if entry["action"] == "updated"}),
        "cells_to_update": len(updates),
        "rows_updated": 0 if dry_run else len({entry["target_row"] for entry in row_audit if entry["action"] == "updated"}),
        "missing_google_headers": [],
        "row_audit": row_audit,
    }
