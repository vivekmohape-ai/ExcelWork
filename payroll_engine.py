from __future__ import annotations

import hashlib
import io
import re
from copy import copy
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from excel_utils import (
    add_new_hire_to_section,
    analyze_template_sheet,
    detect_workbook_month,
    get_days_in_month,
    parse_attendance_input,
    update_active_workbook_month,
)
from identity import (
    IDENTITY_SHEET_NAME,
    ensure_identity_sheet,
    normalize_employee_id,
    normalize_name,
    payroll_mapping_key,
    upsert_identity_register,
)

MONTH_PATTERN = re.compile(
    r"\b"
    r"(january|february|march|april|may|june|july|august|september|"
    r"october|november|december)"
    r"\s*,?\s*"
    r"(20\d{2})\b",
    re.IGNORECASE,
)

ARCHIVE_SHEETS = {
    "AI - TDS",
    IDENTITY_SHEET_NAME,
    "Attendance Mapping",
}


def source_bytes(source: Any) -> bytes:
    if isinstance(source, str):
        with open(
            source,
            "rb",
        ) as handle:
            return handle.read()

    if hasattr(source, "getvalue"):
        return source.getvalue()

    if isinstance(
        source,
        (bytes, bytearray),
    ):
        return bytes(source)

    raise TypeError(
        "Unsupported file source."
    )


def load_workbook(source: Any):
    return openpyxl.load_workbook(
        io.BytesIO(
            source_bytes(source)
        ),
        read_only=False,
        data_only=False,
    )


def sha256_bytes(source: Any) -> str:
    return hashlib.sha256(
        source_bytes(source)
    ).hexdigest()


def _sheet_months(ws) -> List[str]:
    found = []

    max_row = min(
        ws.max_row,
        10,
    )

    max_col = min(
        ws.max_column,
        12,
    )

    for row in ws.iter_rows(
        min_row=1,
        max_row=max_row,
        min_col=1,
        max_col=max_col,
    ):
        for cell in row:
            if not isinstance(
                cell.value,
                str,
            ):
                continue

            for match in MONTH_PATTERN.finditer(
                cell.value
            ):
                found.append(
                    f"{match.group(1).capitalize()} "
                    f"{match.group(2)}"
                )

    return found


def detect_template_month_safe(
    wb,
) -> Optional[str]:
    counts: Dict[str, int] = {}

    for ws in wb.worksheets:
        if ws.title in ARCHIVE_SHEETS:
            continue

        if not analyze_template_sheet(
            ws
        ):
            continue

        for month in _sheet_months(ws):
            counts[month] = (
                counts.get(
                    month,
                    0,
                )
                + 1
            )

    if not counts:
        return detect_workbook_month(
            wb
        )

    return max(
        counts.items(),
        key=lambda item: item[1],
    )[0]


def detect_active_sheet_names(
    wb,
    template_month: Optional[str],
) -> List[str]:
    active = []

    normalized_template = normalize_name(
        template_month or ""
    )

    for ws in wb.worksheets:
        if ws.title in ARCHIVE_SHEETS:
            continue

        sections = analyze_template_sheet(
            ws
        )

        if not sections:
            continue

        months = _sheet_months(
            ws
        )

        normalized_months = {
            normalize_name(month)
            for month in months
        }

        if (
            normalized_template
            and normalized_template
            in normalized_months
        ):
            active.append(
                ws.title
            )
        elif not months:
            active.append(
                ws.title
            )

    return active


def _copy_cell(
    source,
    target,
    formula_target: Optional[str] = None,
):
    value = source.value

    if (
        formula_target is not None
        and isinstance(
            value,
            str,
        )
        and value.startswith("=")
    ):
        try:
            target.value = (
                Translator(
                    value,
                    origin=source.coordinate,
                ).translate_formula(
                    formula_target
                )
            )
        except Exception:
            target.value = value
    else:
        target.value = value

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


def _find_total_row(
    ws,
    section: Dict[str, Any],
) -> Optional[int]:
    data_rows = [
        int(row["row_num"])
        for row in section.get(
            "rows",
            [],
        )
    ]

    if not data_rows:
        return None

    start = max(
        data_rows
    ) + 1

    stop = min(
        ws.max_row,
        start + 5,
    )

    for row_num in range(
        start,
        stop + 1,
    ):
        values = [
            ws.cell(
                row_num,
                col,
            ).value
            for col in range(
                1,
                min(
                    ws.max_column,
                    8,
                )
                + 1,
            )
        ]

        if any(
            isinstance(
                value,
                str,
            )
            and value.strip().upper().startswith(
                "TOTAL"
            )
            for value in values
        ):
            return row_num

    return None


def _formula_columns(
    ws,
    data_rows: Sequence[int],
) -> Dict[int, int]:
    reps = {}

    for row_num in data_rows:
        for col in range(
            1,
            ws.max_column + 1,
        ):
            value = ws.cell(
                row_num,
                col,
            ).value

            if (
                isinstance(
                    value,
                    str,
                )
                and value.startswith("=")
            ):
                reps.setdefault(
                    col,
                    row_num,
                )

    return reps


def repair_data_formulas(
    ws,
    section: Dict[str, Any],
):
    data_rows = [
        int(row["row_num"])
        for row in section.get(
            "rows",
            [],
        )
    ]

    if not data_rows:
        return

    reps = _formula_columns(
        ws,
        data_rows,
    )

    for col, source_row in reps.items():
        source = ws.cell(
            source_row,
            col,
        )

        for row_num in data_rows:
            target = ws.cell(
                row_num,
                col,
            )

            if target.value is None or (
                isinstance(
                    target.value,
                    str,
                )
                and target.value.startswith("=")
            ):
                _copy_cell(
                    source,
                    target,
                    target.coordinate,
                )


def repair_total_row(
    ws,
    section: Dict[str, Any],
):
    data_rows = [
        int(row["row_num"])
        for row in section.get(
            "rows",
            [],
        )
    ]

    if not data_rows:
        return

    total_row = _find_total_row(
        ws,
        section,
    )

    if total_row is None:
        return

    first = min(
        data_rows
    )

    last = max(
        data_rows
    )

    if section["type"] == "employee":
        columns = (
            list(range(6, 15))
            + list(range(16, 23))
        )
    else:
        columns = list(
            range(
                4,
                20,
            )
        )

    for col in columns:
        letter = get_column_letter(
            col
        )

        ws.cell(
            total_row,
            col,
        ).value = (
            f"=SUM({letter}{first}:{letter}{last})"
        )


def repair_payroll_formulas(
    wb,
    template_data: Dict[
        str,
        List[Dict[str, Any]],
    ],
):
    for sheet_name, sections in template_data.items():
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        for section in sections:
            repair_data_formulas(
                ws,
                section,
            )

            repair_total_row(
                ws,
                section,
            )


def update_attendance_rows(
    wb,
    template_data: Dict[
        str,
        List[Dict[str, Any]],
    ],
    mappings_by_key: Dict[
        str,
        Dict[str, Any],
    ],
    total_days: int,
):
    """
    Write FINAL PAYABLE DAYS into the payroll template.

    mapping["days_present"] is kept as the final value after
    leave reconciliation.
    """

    for sheet_name, sections in template_data.items():
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        for section in sections:
            for emp in section.get(
                "rows",
                [],
            ):
                key = payroll_mapping_key(
                    sheet_name,
                    section["type"],
                    int(emp["row_num"]),
                )

                mapping = mappings_by_key.get(
                    key,
                    {},
                )

                present = (
                    float(
                        mapping.get(
                            "final_payable_days"
                        )
                        if mapping.get(
                            "final_payable_days"
                        )
                        is not None
                        else mapping.get(
                            "days_present"
                        )
                        or 0
                    )
                    if mapping.get(
                        "confirmed"
                    )
                    else 0.0
                )

                row_num = int(
                    emp["row_num"]
                )

                if section["type"] == "employee":
                    ws.cell(
                        row_num,
                        5,
                    ).value = total_days

                    ws.cell(
                        row_num,
                        6,
                    ).value = present
                else:
                    ws.cell(
                        row_num,
                        4,
                    ).value = total_days

                    ws.cell(
                        row_num,
                        5,
                    ).value = present


def add_new_hires(
    wb,
    template_data: Dict[
        str,
        List[Dict[str, Any]],
    ],
    new_hires: Iterable[
        Dict[str, Any]
    ],
    total_days: int,
):
    results = []

    for hire in new_hires:
        sheet_name = hire[
            "sheet"
        ]

        section_type = hire[
            "type"
        ]

        sections = template_data.get(
            sheet_name,
            [],
        )

        section = next(
            (
                item
                for item in sections
                if item["type"]
                == section_type
            ),
            None,
        )

        if section is None:
            raise ValueError(
                f"No {section_type} section exists in "
                f"sheet '{sheet_name}'."
            )

        inserted = add_new_hire_to_section(
            wb[sheet_name],
            section,
            {
                "name": hire[
                    "attendance_name"
                ],
                "designation": hire.get(
                    "designation",
                    "",
                ),
                "branch": hire.get(
                    "branch",
                    "",
                ),
                "gross_salary": hire.get(
                    "gross_salary",
                    0,
                ),
            },
            total_days=total_days,
            days_present=float(
                hire.get(
                    "days_present"
                )
                or 0
            ),
        )

        if inserted is None:
            raise ValueError(
                f"No blank row available in "
                f"'{sheet_name}' for "
                f"{hire['attendance_name']}."
            )

        section["rows"].append(
            {
                "row_num": inserted,
                "payroll_code": wb[
                    sheet_name
                ].cell(
                    inserted,
                    1,
                ).value,
                "name": hire[
                    "attendance_name"
                ],
                "designation": hire.get(
                    "designation",
                    "",
                ),
                "branch": hire.get(
                    "branch",
                    "",
                ),
                "gross_salary": hire.get(
                    "gross_salary",
                    0,
                ),
                "total_days": total_days,
                "days_present": hire.get(
                    "days_present",
                    0,
                ),
            }
        )

        results.append(
            {
                "essl_id": normalize_employee_id(
                    hire["essl_id"]
                ),
                "attendance_name": hire[
                    "attendance_name"
                ],
                "sheet_name": sheet_name,
                "section_type": section_type,
                "row_num": inserted,
                "payroll_code": wb[
                    sheet_name
                ].cell(
                    inserted,
                    1,
                ).value,
                "payroll_name": hire[
                    "attendance_name"
                ],
                "match_method": (
                    "new_hire_confirmed"
                ),
                "confirmed": True,
            }
        )

    return results


def write_attendance_mapping_sheet(
    wb,
    mappings: Iterable[
        Dict[str, Any]
    ],
    payroll_month: str,
):
    if "Attendance Mapping" in wb.sheetnames:
        del wb[
            "Attendance Mapping"
        ]

    ws = wb.create_sheet(
        "Attendance Mapping"
    )

    headers = [
        "Payroll Sheet",
        "Section",
        "Payroll Row",
        "Payroll Name",
        "Payroll Code",
        "ESSL Employee ID",
        "Attendance Name",
        "ESSL Payroll Present Days",
        "Approved Leaves",
        "Approved Comp Offs",
        "Unpaid Leaves",
        "Leave Adjustment Applied",
        "Final Payable Days",
        "Leave Match Method",
        "Leave Report Name",
        "Leave Match Score",
        "Match Status",
        "Notes",
        "Payroll Action",
    ]

    ws.append(headers)

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        color="FFFFFF",
        bold=True,
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    for mapping in mappings:
        status = mapping.get(
            "status",
            "",
        )

        confirmed = bool(
            mapping.get(
                "confirmed"
            )
        )

        essl_id = normalize_employee_id(
            mapping.get(
                "essl_id"
            )
        )

        attendance_name = str(
            mapping.get(
                "attendance_name"
            )
            or ""
        ).strip()

        final_days = float(
            mapping.get(
                "final_payable_days"
            )
            if mapping.get(
                "final_payable_days"
            )
            is not None
            else mapping.get(
                "days_present"
            )
            or 0
        )

        if mapping.get(
            "payroll_action"
        ):
            action = mapping[
                "payroll_action"
            ]
        elif confirmed and essl_id:
            action = "UPDATED"
        elif status == "NOT PRESENT":
            action = "SET 0 DAYS"
        else:
            action = "REVIEW"

        ws.append(
            [
                mapping.get(
                    "sheet_name"
                ),
                mapping.get(
                    "section_type"
                ),
                mapping.get(
                    "row_num"
                ),
                mapping.get(
                    "payroll_name"
                ),
                mapping.get(
                    "payroll_code"
                ),
                essl_id or None,
                attendance_name or None,
                float(
                    mapping.get(
                        "payroll_present_days"
                    )
                    or 0
                ),
                float(
                    mapping.get(
                        "approved_leaves"
                    )
                    or 0
                ),
                float(
                    mapping.get(
                        "approved_comp_offs"
                    )
                    or 0
                ),
                float(
                    mapping.get(
                        "unpaid_leaves"
                    )
                    or 0
                ),
                float(
                    mapping.get(
                        "leave_adjustment"
                    )
                    or 0
                ),
                final_days,
                mapping.get(
                    "leave_match_method",
                    "",
                ),
                mapping.get(
                    "leave_report_name",
                    "",
                ),
                round(
                    float(
                        mapping.get(
                            "leave_match_score"
                        )
                        or 0
                    ),
                    3,
                ),
                status,
                (
                    mapping.get(
                        "leave_note"
                    )
                    or mapping.get(
                        "note",
                        "",
                    )
                ),
                action,
            ]
        )

    ws.freeze_panes = "A2"

    widths = [
        20,
        12,
        12,
        28,
        12,
        18,
        24,
        24,
        18,
        20,
        18,
        24,
        20,
        22,
        26,
        16,
        22,
        80,
        18,
    ]

    for index, width in enumerate(
        widths,
        start=1,
    ):
        ws.column_dimensions[
            get_column_letter(index)
        ].width = width

    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False


def validate_attendance_records(
    attendance_records: Sequence[
        Dict[str, Any]
    ],
) -> List[str]:
    issues = []
    seen = {}

    for record in attendance_records:
        emp_id = normalize_employee_id(
            record.get(
                "employee_id"
            )
        )

        name = str(
            record.get(
                "name"
            )
            or ""
        ).strip()

        if not emp_id:
            issues.append(
                f"Attendance row '{name}' has no Employee ID."
            )
            continue

        if not name:
            issues.append(
                f"Attendance ID {emp_id} has no employee name."
            )
            continue

        if emp_id in seen:
            issues.append(
                f"Duplicate ESSL Employee ID {emp_id}: "
                f"'{seen[emp_id]}' and '{name}'."
            )
        else:
            seen[
                emp_id
            ] = name

        payroll_days = float(
            record.get(
                "payroll_present_days"
            )
            if record.get(
                "payroll_present_days"
            )
            is not None
            else record.get(
                "days_present"
            )
            or 0
        )

        if payroll_days < 0:
            issues.append(
                f"{name}: Payroll Present Days cannot be negative."
            )

    return issues


def archive_previous_ai_consultants(
    wb,
    old_month_year: str,
):
    """
    Archive the consultant rows from AI into AI - TDS.

    Historical archive behavior is preserved.
    """

    if (
        "AI" not in wb.sheetnames
        or "AI - TDS" not in wb.sheetnames
    ):
        return

    ai = wb["AI"]
    tds = wb["AI - TDS"]

    sections = analyze_template_sheet(
        ai
    )

    section = next(
        (
            item
            for item in sections
            if item["type"] == "consultant"
            and item.get("rows")
        ),
        None,
    )

    if section is None:
        return

    data_rows = [
        int(row["row_num"])
        for row in section[
            "rows"
        ]
    ]

    total_row = _find_total_row(
        ai,
        section,
    )

    if not data_rows or total_row is None:
        return

    last_row = tds.max_row

    while (
        last_row > 1
        and not any(
            tds.cell(
                last_row,
                col,
            ).value
            for col in range(
                1,
                min(
                    tds.max_column,
                    25,
                )
                + 1,
            )
        )
    ):
        last_row -= 1

    start = last_row + 3

    tds.cell(
        start,
        1,
    ).value = (
        "Consultant Fees for the Month of "
        f"{old_month_year}"
    )

    for offset in range(
        1,
        5,
    ):
        source_row = offset + 1
        target_row = start + offset

        for col in range(
            1,
            min(
                tds.max_column,
                25,
            )
            + 1,
        ):
            _copy_cell(
                tds.cell(
                    source_row,
                    col,
                ),
                tds.cell(
                    target_row,
                    col,
                ),
            )

    target_data_start = start + 5

    for index, row_info in enumerate(
        section["rows"]
    ):
        target_row = (
            target_data_start
            + index
        )

        _copy_cell(
            ai.cell(
                row_info["row_num"],
                1,
            ),
            tds.cell(
                target_row,
                1,
            ),
        )

        for col in range(
            2,
            min(
                ai.max_column,
                tds.max_column,
            )
            + 1,
        ):
            source = ai.cell(
                row_info["row_num"],
                col,
            )

            target = tds.cell(
                target_row,
                col,
            )

            _copy_cell(
                source,
                target,
                target.coordinate,
            )

    target_total = (
        target_data_start
        + len(
            section["rows"]
        )
    )

    for col in range(
        1,
        min(
            ai.max_column,
            tds.max_column,
        )
        + 1,
    ):
        _copy_cell(
            ai.cell(
                total_row,
                col,
            ),
            tds.cell(
                target_total,
                col,
            ),
            tds.cell(
                target_total,
                col,
            ).coordinate,
        )


def archive_already_exists(
    wb,
    month_year: str,
) -> bool:
    if "AI - TDS" not in wb.sheetnames:
        return False

    ws = wb["AI - TDS"]

    needle = normalize_name(
        "Consultant Fees for the Month of "
        f"{month_year}"
    )

    for row in ws.iter_rows(
        min_row=1,
        max_row=min(
            ws.max_row,
            1000,
        ),
    ):
        for cell in row[:3]:
            if (
                isinstance(
                    cell.value,
                    str,
                )
                and normalize_name(
                    cell.value
                ) == needle
            ):
                return True

    return False


def generate_payroll_workbook(
    template_source: Any,
    attendance_source: Any,
    target_month: str,
    mappings_by_key: Dict[
        str,
        Dict[str, Any],
    ],
    new_hires: Optional[
        Iterable[Dict[str, Any]]
    ] = None,
    archive_enabled: bool = True,
) -> Tuple[
    bytes,
    List[Dict[str, Any]],
    List[str],
]:
    """
    Generate payroll from the template.

    The attendance source is parsed for validation only. The actual
    values written into payroll rows come from mappings_by_key, where
    Leave / Comp Off reconciliation has already been applied.
    """

    new_hires = list(
        new_hires or []
    )

    wb = load_workbook(
        template_source
    )

    template_month = detect_template_month_safe(
        wb
    )

    active_sheet_names = detect_active_sheet_names(
        wb,
        template_month,
    )

    template_data = {
        sheet_name: analyze_template_sheet(
            wb[sheet_name]
        )
        for sheet_name in active_sheet_names
    }

    total_days = get_days_in_month(
        target_month
    )

    if (
        archive_enabled
        and template_month
        and normalize_name(
            template_month
        )
        != normalize_name(
            target_month
        )
        and not archive_already_exists(
            wb,
            template_month,
        )
    ):
        archive_previous_ai_consultants(
            wb,
            template_month,
        )

    update_active_workbook_month(
        wb,
        target_month,
        skip_sheets=(
            set(wb.sheetnames)
            - set(active_sheet_names)
        ),
    )

    update_attendance_rows(
        wb,
        template_data,
        mappings_by_key,
        total_days,
    )

    new_hire_results = add_new_hires(
        wb,
        template_data,
        new_hires,
        total_days,
    )

    template_data_after = {
        sheet_name: analyze_template_sheet(
            wb[sheet_name]
        )
        for sheet_name in active_sheet_names
        if sheet_name in wb.sheetnames
    }

    update_attendance_rows(
        wb,
        template_data_after,
        mappings_by_key,
        total_days,
    )

    repair_payroll_formulas(
        wb,
        template_data_after,
    )

    audit_rows = []

    for mapping in mappings_by_key.values():
        audit_rows.append(
            dict(mapping)
        )

    for result in new_hire_results:
        audit_rows.append(
            {
                **result,
                "status": "NEW HIRE",
                "note": (
                    "New payroll row inserted "
                    "and linked to ESSL ID."
                ),
                "payroll_action": "ADDED",
            }
        )

    audit_rows.sort(
        key=lambda row: (
            str(
                row.get(
                    "sheet_name"
                )
                or ""
            ),
            str(
                row.get(
                    "section_type"
                )
                or ""
            ),
            int(
                row.get(
                    "row_num"
                )
                or 0
            ),
        )
    )

    write_attendance_mapping_sheet(
        wb,
        audit_rows,
        target_month,
    )

    ensure_identity_sheet(
        wb
    )

    register_mappings = []

    for mapping in mappings_by_key.values():
        if (
            not mapping.get(
                "confirmed"
            )
            or not mapping.get(
                "essl_id"
            )
        ):
            continue

        register_mappings.append(
            {
                "essl_id": normalize_employee_id(
                    mapping.get(
                        "essl_id"
                    )
                ),
                "attendance_name": mapping.get(
                    "attendance_name",
                    "",
                ),
                "sheet_name": mapping.get(
                    "sheet_name",
                    "",
                ),
                "section_type": mapping.get(
                    "section_type",
                    "",
                ),
                "row_num": mapping.get(
                    "row_num"
                ),
                "payroll_code": mapping.get(
                    "payroll_code"
                ),
                "payroll_name": mapping.get(
                    "payroll_name",
                    "",
                ),
                "match_method": mapping.get(
                    "match_method",
                    "user_confirmed",
                ),
                "confirmed": True,
            }
        )

    register_mappings.extend(
        new_hire_results
    )

    upsert_identity_register(
        wb,
        register_mappings,
        payroll_month=target_month,
    )

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output = io.BytesIO()

    wb.save(output)
    wb.close()

    return (
        output.getvalue(),
        audit_rows,
        active_sheet_names,
    )


def validate_attendance_source(
    attendance_source: Any,
) -> List[str]:
    """
    Optional helper kept for direct testing.
    """
    _, records = parse_attendance_input(
        attendance_source
    )
    return validate_attendance_records(
        records
    )
