
from __future__ import annotations

import calendar
import io
import re
from datetime import datetime, date
from copy import copy
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd
import openpyxl

from openpyxl.formula.translate import Translator


MONTH_PATTERN = re.compile(
    r"\b"
    r"(january|february|march|april|may|june|july|august|september|"
    r"october|november|december)"
    r"\s*,?\s*"
    r"(20\d{2})\b",
    re.IGNORECASE,
)

HEADER_SYNONYMS = {
    "employee_id": {
        "employee id",
        "employee code",
        "emp code",
        "emp. code",
        "code",
        "id",
    },
    "name": {
        "employee name",
        "name of employee",
        "employee",
        "employee name",
        "name",
        "name of consultant",
        "consultant",
        "consultant name",
        "name of freelancer",
        "freelancer",
    },
    "days_present": {
        "days present",
        "no days present",
        "dayspresent",
        "present days",
        "present",
        "attendance days",
        "working days present",
        "no. days present",
    },
    "absent_count": {
        "absent count",
        "absence count",
        "days absent",
        "absent",
        "a",
    },
    "absent_days": {
        "absent days",
        "absence days",
    },
}


def _normalize_header(value: Any) -> str:
    if value is None:
        return ""

    text = str(value).strip().lower()
    text = text.replace("\xa0", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _find_header_columns(row: Iterable[Any]) -> Dict[str, int]:
    normalized = [
        _normalize_header(value)
        for value in row
    ]

    columns: Dict[str, int] = {}

    for idx, value in enumerate(normalized):
        if not value:
            continue

        for field_name, variants in HEADER_SYNONYMS.items():
            if field_name in columns:
                continue

            if value in variants:
                columns[field_name] = idx
                break

            if field_name == "days_present" and (
                "present" in value
                or (
                    "days" in value
                    and "present" in value
                )
            ):
                columns[field_name] = idx
                break

    return columns


def _extract_month_year_from_value(value: Any) -> Optional[str]:
    if value is None:
        return None

    match = MONTH_PATTERN.search(
        str(value)
    )

    if not match:
        return None

    return (
        f"{match.group(1).capitalize()} "
        f"{match.group(2)}"
    )


def infer_month_year(
    file_content: Any = None,
    filename: Optional[str] = None,
) -> Optional[str]:
    """
    Infer target month/year from the uploaded attendance file
    or its filename.
    """

    if filename:
        month = _extract_month_year_from_value(
            filename.replace("_", " ")
        )
        if month:
            return month

    try:
        if isinstance(file_content, (bytes, bytearray)):
            source = io.BytesIO(
                file_content
            )
        elif hasattr(
            file_content,
            "seek",
        ):
            file_content.seek(0)
            source = file_content
        else:
            source = file_content

        df = pd.read_excel(
            source,
            header=None,
            nrows=20,
        )

        for value in df.to_numpy().flatten():
            month = _extract_month_year_from_value(
                value
            )
            if month:
                return month

    except Exception:
        pass

    return None



def _normalize_employee_id_local(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip().replace(",", "")
    if not text:
        return ""
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except (TypeError, ValueError):
        pass
    return text


def _read_dataframe_source(source: Any, filename: Optional[str] = None) -> pd.DataFrame:
    """Read an uploaded attendance/punch source as a dataframe."""
    if isinstance(source, pd.DataFrame):
        return source.copy()

    if hasattr(source, "seek"):
        try:
            source.seek(0)
        except Exception:
            pass

    name = (filename or getattr(source, "name", "") or "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(source)
    return pd.read_excel(source, header=None)


def _normalize_column_label(value: Any) -> str:
    text = _normalize_header(value)
    return text.replace(" ", "_")


def _find_named_column(columns: Iterable[Any], aliases: Iterable[str]) -> Optional[int]:
    normalized = [_normalize_column_label(c) for c in columns]
    alias_set = {_normalize_column_label(a) for a in aliases}
    for idx, value in enumerate(normalized):
        if value in alias_set:
            return idx
    for idx, value in enumerate(normalized):
        if any(alias in value for alias in alias_set):
            return idx
    return None


def _coerce_date_key(value: Any) -> Optional[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    try:
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=False)
    except Exception:
        return None
    if pd.isna(parsed):
        return None
    return parsed.date().isoformat()


def _has_punch_value(value: Any) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass
    text = str(value).strip().lower()
    if not text or text in {"nan", "nat", "none", "null", "-", "--"}:
        return False
    return True


def _looks_like_datetime_punch(value: Any) -> bool:
    """Return True when a single timestamp cell contains an actual time-of-day."""
    if value is None:
        return False
    if isinstance(value, datetime):
        return not (value.hour == 0 and value.minute == 0 and value.second == 0)
    text = str(value).strip()
    if not text:
        return False
    return bool(re.search(r"\d{1,2}:\d{2}", text))


def _extract_punch_days(source: Any, filename: Optional[str] = None) -> Dict[str, set[str]]:
    """
    Extract dates on which an employee has any punch evidence.

    Supported layouts include:
      Employee ID | Date | In Time | Out Time
      Employee Code | Attendance Date | Punch In | Punch Out
      Employee ID | DateTime | Punch Type | Punch Time
      one row per punch event with Employee ID + timestamp

    Business rule:
      if either punch in OR punch out exists for a date, that date is PRESENT.
    """
    if source is None:
        return {}

    try:
        df = _read_dataframe_source(source, filename)
    except Exception:
        return {}

    if df.empty:
        return {}

    # Locate a header row for files that are exported with title rows above the data.
    header_idx = 0
    for ridx in range(min(len(df), 25)):
        row = df.iloc[ridx].tolist()
        norm = [_normalize_column_label(v) for v in row]
        has_id = any(v in {"employee_id", "employee_code", "emp_code", "id", "code"} for v in norm)
        has_date = any("date" in v or "time" in v or "timestamp" in v or "datetime" in v for v in norm)
        if has_id and has_date:
            header_idx = ridx
            break

    if header_idx != 0 or not all(isinstance(c, str) for c in df.columns):
        df = df.iloc[header_idx:].copy()
        df.columns = [str(v).strip() for v in df.iloc[0].tolist()]
        df = df.iloc[1:].reset_index(drop=True)

    columns = list(df.columns)
    id_idx = _find_named_column(
        columns,
        ["Employee ID", "Employee Code", "Emp Code", "ID", "Code", "User ID", "User Code"],
    )
    name_idx = _find_named_column(columns, ["Employee Name", "Name", "User Name"])
    date_idx = _find_named_column(
        columns,
        ["Date", "Attendance Date", "Punch Date", "Transaction Date", "Log Date", "Date Time", "Datetime", "Timestamp"],
    )
    in_idx = _find_named_column(
        columns,
        ["In Time", "Punch In", "PunchIn", "First In", "Check In", "Check-In", "Login", "InTime"],
    )
    out_idx = _find_named_column(
        columns,
        ["Out Time", "Punch Out", "PunchOut", "Last Out", "Check Out", "Check-Out", "Logout", "OutTime"],
    )
    punch_time_idx = _find_named_column(
        columns,
        ["Punch Time", "PunchTime", "Event Time", "Transaction Time", "Time"],
    )
    direction_idx = _find_named_column(
        columns,
        ["Punch Type", "Punch Direction", "Direction", "In Out", "Type"],
    )

    if id_idx is None or date_idx is None:
        return {}

    by_id: Dict[str, set[str]] = {}
    for row in df.itertuples(index=False, name=None):
        if id_idx >= len(row) or date_idx >= len(row):
            continue
        emp_id = _normalize_employee_id_local(row[id_idx])
        if not emp_id:
            continue

        raw_date = row[date_idx]
        date_key = _coerce_date_key(raw_date)
        if not date_key:
            continue

        # Any In Time OR Out Time means the employee was present that day.
        punch_found = False
        if in_idx is not None and in_idx < len(row):
            punch_found = punch_found or _has_punch_value(row[in_idx])
        if out_idx is not None and out_idx < len(row):
            punch_found = punch_found or _has_punch_value(row[out_idx])
        if punch_time_idx is not None and punch_time_idx < len(row):
            punch_found = punch_found or _has_punch_value(row[punch_time_idx])
        if direction_idx is not None and direction_idx < len(row):
            direction = str(row[direction_idx]).strip().lower()
            punch_found = punch_found or direction in {
                "in", "out", "punch in", "punch out",
                "check in", "check out", "login", "logout",
            }

        # For one-row-per-event exports, a timestamp in the date/datetime field
        # is itself punch evidence when explicit punch columns are absent.
        if not punch_found and punch_time_idx is None and in_idx is None and out_idx is None:
            punch_found = _looks_like_datetime_punch(row[date_idx])

        if punch_found:
            by_id.setdefault(emp_id, set()).add(date_key)

    return by_id


def apply_punch_evidence(
    attendance_records: List[Dict[str, Any]],
    punch_source: Any,
    punch_filename: Optional[str] = None,
    *,
    punch_log_is_authoritative: bool = True,
) -> Dict[str, set[str]]:
    """Apply the missing-punch rule to attendance records in place.

    For a full monthly punch export, the unique dates containing at least
    one In or Out punch are authoritative: one punch on a date = 1 present day.
    When a detailed daily attendance workbook also provides exact present dates,
    those dates are unioned with punch dates so a one-sided punch can never turn
    into an absence.
    """
    punch_days = _extract_punch_days(punch_source, punch_filename)
    if not punch_days:
        return {}

    for record in attendance_records:
        emp_id = _normalize_employee_id_local(record.get("employee_id"))
        if not emp_id:
            continue

        dates = punch_days.get(emp_id, set())
        detailed_present_dates = {
            str(d) for d in (record.get("present_dates") or []) if d
        }
        combined_dates = detailed_present_dates | set(dates)

        if not combined_dates:
            record["punch_present_days"] = 0
            record["punch_present_dates"] = ""
            record["punch_rule_applied"] = False
            continue

        summary_days = float(record.get("days_present") or 0)
        punch_day_count = float(len(dates))

        if detailed_present_dates:
            # Exact union is possible because the detailed daily source tells us
            # which dates were already counted. One-sided punches turn an A into P.
            corrected_days = float(len(combined_dates))
            # Preserve half-day weighting only when the detailed source explicitly
            # contains half-day statuses and there is no punch on those dates.
            half_dates = {
                str(d) for d in (record.get("half_present_dates") or []) if d
            }
            corrected_days -= 0.5 * len(half_dates - set(dates))
            source_label = "daily attendance + punch evidence"
        elif punch_log_is_authoritative:
            # With only an aggregate summary there is no way to know date overlap.
            # A full raw punch export is therefore authoritative when it is at
            # least as complete as the summary. If it has fewer punch dates, keep
            # the summary value and flag the discrepancy rather than risking an
            # underpayment caused by an incomplete export.
            if punch_day_count >= summary_days:
                corrected_days = punch_day_count
                source_label = "full raw punch log"
            else:
                corrected_days = summary_days
                source_label = "attendance summary retained because punch log is lower"
                record["punch_discrepancy"] = True
                record["punch_discrepancy_note"] = (
                    f"Punch log contains {int(punch_day_count)} unique punch date(s), "
                    f"while the attendance summary contains {summary_days:g} present day(s). "
                    "Check that the raw punch export is complete."
                )
        else:
            # Do not guess. An exception-only punch file cannot safely be unioned
            # with an aggregate monthly total because overlap is unknown.
            corrected_days = summary_days
            source_label = "attendance summary only"

        if corrected_days != summary_days:
            record["days_present_before_punch_rule"] = summary_days
            record["days_present"] = corrected_days
            record["punch_rule_applied"] = True
            record["punch_rule_note"] = (
                f"Present days recalculated using {source_label}. "
                "Any date with either Punch In or Punch Out counts as 1 present day."
            )
        else:
            record["punch_rule_applied"] = False

        record["punch_present_days"] = int(punch_day_count)
        record["punch_present_dates"] = ", ".join(sorted(dates))

    return punch_days


def parse_attendance_input(
    file_content: Any,
    filename: Optional[str] = None,
    punch_source: Any = None,
    punch_filename: Optional[str] = None,
    *,
    punch_log_is_authoritative: bool = True,
):
    """
    Parse both supported attendance workbook formats.

    Supported format A:

        Employee ID
        Employee Name
        Days Present
        Absent Count

    Supported format B:

        Sr No.
        Employee Code
        Employee Name
        1
        2
        ...
        31

    Returns:

        month_year,
        records

    Every record contains:

        employee_id
        name
        days_present
        absent_count
        absent_days
    """

    month_year = infer_month_year(
        file_content=file_content,
        filename=filename,
    )

    if hasattr(
        file_content,
        "seek",
    ):
        file_content.seek(0)

    df = pd.read_excel(
        file_content,
        header=None,
    )

    if df.empty:
        raise ValueError(
            "Attendance workbook is empty."
        )

    header_row_idx = None
    header_columns: Dict[str, int] = {}

    for row_idx in range(
        min(len(df), 30)
    ):
        columns = _find_header_columns(
            df.iloc[row_idx].tolist()
        )

        if (
            "name" in columns
            and (
                "days_present" in columns
                or "employee_id" in columns
            )
        ):
            header_row_idx = row_idx
            header_columns = columns
            break

    if header_row_idx is None:
        raise ValueError(
            "Could not identify the attendance header row."
        )

    # ---------------------------------------------------------
    # Identify ID / name / days present columns
    # ---------------------------------------------------------

    name_col = header_columns.get(
        "name"
    )

    id_col = header_columns.get(
        "employee_id"
    )

    days_present_col = header_columns.get(
        "days_present"
    )

    absent_count_col = header_columns.get(
        "absent_count"
    )

    absent_days_col = header_columns.get(
        "absent_days"
    )

    if name_col is None:
        raise ValueError(
            "Attendance workbook does not contain an employee name column."
        )

    # ---------------------------------------------------------
    # Determine daily date/status columns
    # ---------------------------------------------------------

    daily_columns: Dict[int, int] = {}

    header_values = (
        df.iloc[header_row_idx]
        .tolist()
    )

    for col_idx, value in enumerate(
        header_values
    ):
        if isinstance(
            value,
            (int, float),
        ):
            day = int(value)

            if 1 <= day <= 31:
                daily_columns[day] = col_idx
            continue

        text = str(
            value
        ).strip()

        if re.fullmatch(
            r"\d{1,2}",
            text,
        ):
            day = int(text)

            if 1 <= day <= 31:
                daily_columns[day] = col_idx

    # ---------------------------------------------------------
    # Parse rows
    # ---------------------------------------------------------

    records: List[Dict[str, Any]] = []

    for row_idx in range(
        header_row_idx + 1,
        len(df),
    ):
        row = df.iloc[row_idx]

        raw_name = (
            row.iloc[name_col]
            if name_col < len(row)
            else None
        )

        if pd.isna(raw_name):
            continue

        name = str(
            raw_name
        ).strip()

        if not name:
            continue

        # Skip subtotal / footer rows.
        if name.upper() in {
            "TOTAL",
            "DEPARTMENT DEFAULT",
        }:
            continue

        # -----------------------------------------------------
        # Employee ID
        # -----------------------------------------------------

        if id_col is not None:
            raw_id = (
                row.iloc[id_col]
                if id_col < len(row)
                else None
            )

            if pd.isna(raw_id):
                employee_id = ""
            else:
                text_id = str(
                    raw_id
                ).strip()

                try:
                    numeric_id = float(
                        text_id
                    )

                    if numeric_id.is_integer():
                        employee_id = str(
                            int(numeric_id)
                        )
                    else:
                        employee_id = text_id

                except (
                    TypeError,
                    ValueError,
                ):
                    employee_id = text_id

        else:
            # Old attendance format:
            # employee code is normally column 2.
            employee_id = ""

        if not employee_id:
            # Try the second column as an employee code.
            if len(row) > 1:
                raw_id = row.iloc[1]

                if pd.notna(raw_id):
                    text_id = str(
                        raw_id
                    ).strip()

                    try:
                        numeric_id = float(
                            text_id
                        )

                        if numeric_id.is_integer():
                            employee_id = str(
                                int(numeric_id)
                            )
                        else:
                            employee_id = text_id

                    except (
                        TypeError,
                        ValueError,
                    ):
                        employee_id = text_id

        # -----------------------------------------------------
        # Days present
        # -----------------------------------------------------

        days_present: Optional[float] = None

        if days_present_col is not None:
            raw_present = (
                row.iloc[
                    days_present_col
                ]
                if days_present_col < len(row)
                else None
            )

            if pd.notna(raw_present):
                try:
                    days_present = float(
                        raw_present
                    )
                except (
                    TypeError,
                    ValueError,
                ):
                    days_present = None

        # -----------------------------------------------------
        # Absent count
        # -----------------------------------------------------

        absent_count: Optional[float] = None

        if absent_count_col is not None:
            raw_absent = (
                row.iloc[
                    absent_count_col
                ]
                if absent_count_col < len(row)
                else None
            )

            if pd.notna(raw_absent):
                try:
                    absent_count = float(
                        raw_absent
                    )
                except (
                    TypeError,
                    ValueError,
                ):
                    absent_count = None

        # -----------------------------------------------------
        # Absent days
        # -----------------------------------------------------

        absent_days = ""

        if absent_days_col is not None:
            raw_absent_days = (
                row.iloc[
                    absent_days_col
                ]
                if absent_days_col < len(row)
                else None
            )

            if pd.notna(
                raw_absent_days
            ):
                absent_days = str(
                    raw_absent_days
                ).strip()

        # -----------------------------------------------------
        # Daily status details and fallback
        # -----------------------------------------------------

        present_dates: List[str] = []
        half_present_dates: List[str] = []

        for day, col_idx in sorted(
            daily_columns.items()
        ):
            if col_idx >= len(row):
                continue

            raw_status = row.iloc[col_idx]
            if pd.isna(raw_status):
                continue

            status = str(raw_status).strip().upper()
            if month_year:
                month_match = MONTH_PATTERN.search(str(month_year))
                if month_match:
                    month_number = list(calendar.month_name).index(month_match.group(1).capitalize())
                    date_key = f"{int(month_match.group(2)):04d}-{month_number:02d}-{int(day):02d}"
                else:
                    date_key = str(day)
            else:
                date_key = str(day)

            if status in {"P", "PRESENT"}:
                present_dates.append(date_key)
            elif status in {"½P", "1/2P", "0.5P", "HALF DAY", "HALF"}:
                half_present_dates.append(date_key)

        if days_present is None:
            present_total = float(len(present_dates)) + (0.5 * len(half_present_dates))

            for day, col_idx in sorted(
                daily_columns.items()
            ):
                if col_idx >= len(row):
                    continue

                raw_status = row.iloc[
                    col_idx
                ]

                if pd.isna(raw_status):
                    continue

                status = str(
                    raw_status
                ).strip().upper()

                if status in {
                    "P",
                    "PRESENT",
                }:
                    present_total += 1.0

                elif status in {
                    "½P",
                    "1/2P",
                    "0.5P",
                    "HALF DAY",
                    "HALF",
                }:
                    present_total += 0.5

            days_present = present_total

        if absent_count is None:
            if daily_columns:
                absent_count = sum(
                    1
                    for _, col_idx in sorted(
                        daily_columns.items()
                    )
                    if (
                        col_idx < len(row)
                        and pd.notna(
                            row.iloc[col_idx]
                        )
                        and str(
                            row.iloc[col_idx]
                        ).strip().upper()
                        == "A"
                    )
                )
            else:
                absent_count = 0.0

        records.append(
            {
                "employee_id": employee_id,
                "name": name,
                "days_present": float(
                    days_present or 0
                ),
                "absent_count": float(
                    absent_count or 0
                ),
                "absent_days": absent_days,
                "present_dates": present_dates,
                "half_present_dates": half_present_dates,
            }
        )

    if not records:
        raise ValueError(
            "No attendance records were found."
        )

    # Correct the monthly total using raw punch evidence when supplied.
    # A single In or Out punch on a date counts as present for that date.
    apply_punch_evidence(
        records,
        punch_source,
        punch_filename=punch_filename,
        punch_log_is_authoritative=punch_log_is_authoritative,
    )

    return month_year, records


def detect_workbook_month(
    wb,
) -> Optional[str]:
    """
    Detect an existing payroll month/year from workbook text.
    Searches titles and early rows.
    """

    for ws in wb.worksheets:

        max_scan_row = min(
            ws.max_row,
            8,
        )

        max_scan_col = min(
            ws.max_column,
            8,
        )

        for r in range(
            1,
            max_scan_row + 1,
        ):
            for c in range(
                1,
                max_scan_col + 1,
            ):
                value = ws.cell(
                    r,
                    c,
                ).value

                month = (
                    _extract_month_year_from_value(
                        value
                    )
                )

                if month:
                    return month

    return None


def get_days_in_month(
    month_year_str: str,
) -> int:
    """
    Return calendar days for a month/year string.
    """

    if not month_year_str:
        return 30

    match = MONTH_PATTERN.search(
        str(month_year_str)
    )

    if not match:
        return 30

    month_name = (
        match.group(1).lower()
    )

    year = int(
        match.group(2)
    )

    month_number = list(
        calendar.month_name
    ).index(
        month_name.capitalize()
    )

    return calendar.monthrange(
        year,
        month_number,
    )[1]


def update_headers_in_sheet(
    sheet,
    old_month_str: Optional[str],
    new_month_str: str,
):
    """
    Compatibility function.

    Replaces the old month/year wherever it occurs.
    """

    if old_month_str:
        pattern = re.compile(
            re.escape(old_month_str),
            re.IGNORECASE,
        )

        alternate_old = (
            old_month_str.replace(
                " ",
                ", ",
            )
        )

        alternate_new = (
            new_month_str.replace(
                " ",
                ", ",
            )
        )

        alternate_pattern = re.compile(
            re.escape(alternate_old),
            re.IGNORECASE,
        )

    else:
        pattern = None
        alternate_pattern = None
        alternate_new = new_month_str

    for row in sheet.iter_rows():

        for cell in row:

            value = cell.value

            if not isinstance(
                value,
                str,
            ):
                continue

            new_value = value

            if pattern:
                new_value = pattern.sub(
                    new_month_str,
                    new_value,
                )

            if alternate_pattern:
                new_value = alternate_pattern.sub(
                    alternate_new,
                    new_value,
                )

            cell.value = new_value


def update_active_workbook_month(
    wb,
    new_month_str: str,
    skip_sheets: Optional[
        Iterable[str]
    ] = None,
):
    """
    Replace month/year strings in active payroll sheets.

    Historical log sheets such as AI - TDS should normally be skipped.
    """

    skip = {
        str(name)
        for name in (
            skip_sheets or []
        )
    }

    for ws in wb.worksheets:

        if ws.title in skip:
            continue

        for row in ws.iter_rows():

            for cell in row:

                if not isinstance(
                    cell.value,
                    str,
                ):
                    continue

                original = cell.value

                def replace_match(
                    match,
                ):
                    return new_month_str

                cell.value = MONTH_PATTERN.sub(
                    replace_match,
                    original,
                )


def analyze_template_sheet(
    sheet,
) -> List[Dict[str, Any]]:
    """
    Identify employee/consultant sections in one payroll sheet.

    Returns sections with:

        type
        header_row
        start_row
        end_row
        name_col
        rows

    Each row includes:

        row_num
        payroll_code
        name
        designation
        branch
        gross_salary
        total_days
        days_present
    """

    rows = list(
        sheet.iter_rows(
            values_only=True
        )
    )

    section_headers: List[
        tuple[int, str, int]
    ] = []

    for r_idx, row in enumerate(
        rows,
        start=1,
    ):

        normalized = [
            _normalize_header(value)
            for value in row
        ]

        employee_header = None
        consultant_header = None

        for idx, value in enumerate(
            normalized
        ):

            if value in {
                "name of employee",
                "employee name",
            }:
                employee_header = idx

            elif value in {
                "name of consultant",
                "name of freelancer",
            }:
                consultant_header = idx

        if employee_header is not None:
            section_headers.append(
                (
                    r_idx,
                    "employee",
                    employee_header + 1,
                )
            )

        elif consultant_header is not None:
            section_headers.append(
                (
                    r_idx,
                    "consultant",
                    consultant_header + 1,
                )
            )

    sections: List[
        Dict[str, Any]
    ] = []

    for idx, (
        header_row,
        section_type,
        name_col,
    ) in enumerate(
        section_headers
    ):

        next_header_row = (
            section_headers[idx + 1][0]
            if idx + 1 < len(section_headers)
            else len(rows) + 1
        )

        section = {
            "type": section_type,
            "header_row": header_row,
            "name_col": name_col,
            "start_row": header_row + 1,
            "end_row": next_header_row - 1,
            "rows": [],
        }

        found_total = False

        for row_num in range(
            section["start_row"],
            section["end_row"] + 1,
        ):

            row = rows[
                row_num - 1
            ]

            values_upper = [
                str(value).strip().upper()
                for value in row
                if value is not None
            ]

            is_total = (
                any(
                    value == "TOTAL"
                    or value.startswith(
                        "TOTAL "
                    )
                    for value in values_upper
                )
            )

            if is_total:
                section["end_row"] = (
                    row_num - 1
                )
                found_total = True
                break

            if (
                name_col > len(row)
                or row[name_col - 1] is None
            ):
                continue

            name = str(
                row[name_col - 1]
            ).strip()

            if not name:
                continue

            if name.upper() in {
                "TOTAL",
                "NAME OF EMPLOYEE",
                "NAME OF CONSULTANT",
            }:
                continue

            payroll_code = (
                row[0]
                if len(row) >= 1
                else None
            )

            designation = (
                row[2]
                if len(row) >= 3
                else None
            )

            branch = (
                row[3]
                if len(row) >= 4
                else None
            )

            if section_type == "employee":
                total_days = (
                    row[4]
                    if len(row) >= 5
                    else None
                )

                days_present = (
                    row[5]
                    if len(row) >= 6
                    else None
                )

                gross_salary = (
                    row[6]
                    if len(row) >= 7
                    else 0
                )

            else:
                total_days = (
                    row[3]
                    if len(row) >= 4
                    else None
                )

                days_present = (
                    row[4]
                    if len(row) >= 5
                    else None
                )

                gross_salary = (
                    row[5]
                    if len(row) >= 6
                    else 0
                )

            section["rows"].append(
                {
                    "row_num": row_num,
                    "payroll_code": payroll_code,
                    "name": name,
                    "designation": designation,
                    "branch": branch,
                    "gross_salary": gross_salary,
                    "total_days": total_days,
                    "days_present": days_present,
                }
            )

        if (
            not found_total
            and section["end_row"]
            >= len(rows)
        ):
            section["end_row"] = len(
                rows
            )

        sections.append(
            section
        )

    return sections


def _copy_row_with_translated_formulas(
    sheet,
    source_row: int,
    target_row: int,
):
    """
    Copy formatting, formulas, and values from source row to target row.

    Relative references in formulas are translated to the target row.
    """

    max_column = sheet.max_column

    for col_idx in range(
        1,
        max_column + 1,
    ):

        source = sheet.cell(
            source_row,
            col_idx,
        )

        target = sheet.cell(
            target_row,
            col_idx,
        )

        if (
            isinstance(
                source.value,
                str,
            )
            and source.value.startswith("=")
        ):

            try:
                target.value = (
                    Translator(
                        source.value,
                        origin=source.coordinate,
                    ).translate_formula(
                        target.coordinate
                    )
                )
            except Exception:
                target.value = source.value

        else:
            target.value = source.value

        if source.has_style:
            target.font = copy(
                source.font
            )
            target.fill = copy(
                source.fill
            )
            target.border = copy(
                source.border
            )
            target.alignment = copy(
                source.alignment
            )
            target.number_format = (
                source.number_format
            )
            target.protection = copy(
                source.protection
            )

        if source.hyperlink:
            target._hyperlink = copy(
                source.hyperlink
            )

        if source.comment:
            target.comment = copy(
                source.comment
            )


def _find_last_data_row(
    section: Dict[str, Any],
) -> Optional[int]:
    rows = section.get(
        "rows",
        [],
    )

    if not rows:
        return None

    return max(
        int(row["row_num"])
        for row in rows
    )


def add_new_hire_to_section(
    sheet,
    section: Dict[str, Any],
    hire: Dict[str, Any],
    total_days: int,
    days_present: float,
) -> Optional[int]:
    """
    Insert a new hire into the first blank row in the section.

    Returns the inserted row number, or None when no blank row exists.
    """

    start_row = int(
        section["start_row"]
    )

    end_row = int(
        section["end_row"]
    )

    name_col = int(
        section["name_col"]
    )

    empty_row = None

    for row_num in range(
        start_row,
        end_row + 1,
    ):

        name_value = sheet.cell(
            row_num,
            name_col,
        ).value

        if (
            name_value is None
            or not str(name_value).strip()
        ):
            empty_row = row_num
            break

    if empty_row is None:
        return None

    source_row = _find_last_data_row(
        section
    )

    if source_row is not None:
        _copy_row_with_translated_formulas(
            sheet,
            source_row,
            empty_row,
        )

    previous_sr = sheet.cell(
        empty_row - 1,
        1,
    ).value

    try:
        sr_no = int(
            float(
                previous_sr
            )
        ) + 1
    except (
        TypeError,
        ValueError,
    ):
        sr_no = ""

    sheet.cell(
        empty_row,
        1,
    ).value = sr_no

    sheet.cell(
        empty_row,
        name_col,
    ).value = hire["name"]

    # Common designation / branch positions.
    sheet.cell(
        empty_row,
        3,
    ).value = hire.get(
        "designation",
        "",
    )

    sheet.cell(
        empty_row,
        4,
    ).value = hire.get(
        "branch",
        "",
    )

    if section["type"] == "employee":

        sheet.cell(
            empty_row,
            5,
        ).value = total_days

        sheet.cell(
            empty_row,
            6,
        ).value = days_present

        sheet.cell(
            empty_row,
            7,
        ).value = hire.get(
            "gross_salary",
            0,
        )

        # Clear Loan / Advance
        sheet.cell(
            empty_row,
            19,
        ).value = None

    else:

        sheet.cell(
            empty_row,
            4,
        ).value = total_days

        sheet.cell(
            empty_row,
            5,
        ).value = days_present

        sheet.cell(
            empty_row,
            6,
        ).value = hire.get(
            "gross_salary",
            0,
        )

        # Clear Loan / Advance
        sheet.cell(
            empty_row,
            15,
        ).value = None

    return empty_row


def archive_previous_month_consultants(
    wb,
    old_month_year_str: str,
):
    """
    Copy the consultant section from AI into AI - TDS.

    This should be called only when moving from one payroll month
    to another and only when archival is intentionally enabled.
    """

    if (
        "AI" not in wb.sheetnames
        or "AI - TDS" not in wb.sheetnames
    ):
        return

    ai_sheet = wb["AI"]
    tds_sheet = wb["AI - TDS"]

    ai_sections = analyze_template_sheet(
        ai_sheet
    )

    consultant_section = next(
        (
            section
            for section in ai_sections
            if section["type"] == "consultant"
            and section.get("rows")
        ),
        None,
    )

    if consultant_section is None:
        return

    last_row = tds_sheet.max_row

    while (
        last_row > 1
        and not any(
            tds_sheet.cell(
                last_row,
                col_idx,
            ).value
            for col_idx in range(
                1,
                min(
                    tds_sheet.max_column,
                    20,
                )
                + 1,
            )
        )
    ):
        last_row -= 1

    start_append_row = (
        last_row + 3
    )

    tds_sheet.cell(
        start_append_row,
        1,
    ).value = (
        "Consultant Fees for the Month of "
        f"{old_month_year_str}"
    )

    # Copy the top 4 header rows.
    for offset in range(
        1,
        5,
    ):
        source_row = 1 + offset
        target_row = (
            start_append_row
            + offset
        )

        for col_idx in range(
            1,
            min(
                tds_sheet.max_column,
                25,
            )
            + 1,
        ):

            source = tds_sheet.cell(
                source_row,
                col_idx,
            )

            target = tds_sheet.cell(
                target_row,
                col_idx,
            )

            target.value = source.value

            if source.has_style:
                target.font = copy(
                    source.font
                )
                target.fill = copy(
                    source.fill
                )
                target.border = copy(
                    source.border
                )
                target.alignment = copy(
                    source.alignment
                )
                target.number_format = (
                    source.number_format
                )

    target_data_start = (
        start_append_row + 5
    )

    source_rows = consultant_section[
        "rows"
    ]

    for index, source_info in enumerate(
        source_rows
    ):
        source_row = int(
            source_info["row_num"]
        )

        target_row = (
            target_data_start
            + index
        )

        _copy_row_with_translated_formulas(
            ai_sheet,
            source_row,
            target_row,
        )

    target_total_row = (
        target_data_start
        + len(source_rows)
    )

    source_total_row = (
        consultant_section[
            "end_row"
        ]
        + 1
    )

    if source_total_row <= ai_sheet.max_row:
        _copy_row_with_translated_formulas(
            ai_sheet,
            source_total_row,
            target_total_row,
        )
