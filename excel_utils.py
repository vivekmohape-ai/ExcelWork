
from __future__ import annotations

import calendar
import io
import re
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


def parse_attendance_input(
    file_content: Any,
    filename: Optional[str] = None,
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
        # Daily status fallback
        # -----------------------------------------------------

        if days_present is None:
            present_total = 0.0

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
            }
        )

    if not records:
        raise ValueError(
            "No attendance records were found."
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
