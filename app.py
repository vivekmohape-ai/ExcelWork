
from __future__ import annotations

import io
import re
import difflib
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import streamlit as st

from excel_utils import (
    add_new_hire_to_section,
    analyze_template_sheet,
    detect_workbook_month,
    get_days_in_month,
    parse_attendance_input,
    update_active_workbook_month,
    archive_previous_month_consultants,
)
from identity import (
    IDENTITY_SHEET_NAME,
    STATUS_CONFLICT,
    STATUS_MISSING_ATTENDANCE,
    STATUS_NOT_PRESENT,
    STATUS_REGISTERED,
    STATUS_REGISTERED_NAME_VARIANT,
    STATUS_REVIEW,
    build_attendance_indexes,
    ensure_identity_sheet,
    load_identity_register,
    normalize_employee_id,
    normalize_name,
    payroll_mapping_key,
    resolve_template_rows,
    upsert_identity_register,
    validate_mapping_conflicts,
)


# ---------------------------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Payroll Automator",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------------------------

DEFAULTS = {
    "attendance_records": [],
    "template_data": {},
    "template_month": None,
    "target_month": None,
    "mapping_state": {},
    "new_employees": [],
    "generated_file": None,
    "download_filename": None,
    "attendance_key": None,
    "template_key": None,
    "identity_loaded": False,
    "identity_register_count": 0,
    "identity_conflicts": {},
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ---------------------------------------------------------------------------
# STYLE
# ---------------------------------------------------------------------------

st.markdown(
    """
<style>
.header-gradient {
    background: linear-gradient(135deg, #4f46e5 0%, #06b6d4 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-weight: 700;
    font-size: 2.8rem;
    margin-bottom: 0.5rem;
}

.subheader-text {
    color: #6b7280;
    font-size: 1.05rem;
    margin-bottom: 1.5rem;
}

.status-card {
    padding: 0.75rem 1rem;
    border-radius: 8px;
    border: 1px solid #e5e7eb;
    background: #f9fafb;
    margin-bottom: 0.5rem;
}

.review-card {
    padding: 1rem;
    border-radius: 10px;
    border: 1px solid #f59e0b;
    background: #fffbeb;
    margin-bottom: 0.75rem;
}

.ok-card {
    padding: 1rem;
    border-radius: 10px;
    border: 1px solid #10b981;
    background: #ecfdf5;
    margin-bottom: 0.75rem;
}
</style>
""",
    unsafe_allow_html=True,
)


st.markdown(
    "<h1 class='header-gradient'>Payroll Automator</h1>",
    unsafe_allow_html=True,
)

st.markdown(
    "<p class='subheader-text'>"
    "Monthly payroll processing with persistent ESSL identity matching."
    "</p>",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

NOT_PRESENT_VALUE = "__NOT_PRESENT__"


def _source_bytes(source: Any) -> bytes:
    if isinstance(source, str):
        with open(source, "rb") as handle:
            return handle.read()

    if hasattr(source, "getvalue"):
        return source.getvalue()

    if isinstance(source, (bytes, bytearray)):
        return bytes(source)

    raise TypeError(
        "Unsupported file source."
    )


def _load_workbook_from_source(
    source: Any,
    *,
    read_only: bool,
):
    data = _source_bytes(source)

    return openpyxl.load_workbook(
        io.BytesIO(data),
        read_only=read_only,
        data_only=False,
    )


def _template_sections(
    wb,
) -> Dict[str, List[Dict[str, Any]]]:
    result: Dict[str, List[Dict[str, Any]]] = {}

    for sheet_name in wb.sheetnames:

        if sheet_name == "AI - TDS":
            continue

        if sheet_name == IDENTITY_SHEET_NAME:
            continue

        sections = analyze_template_sheet(
            wb[sheet_name]
        )

        if sections:
            result[sheet_name] = sections

    return result


def _mapping_option_label(
    record: Dict[str, Any],
) -> str:
    employee_id = normalize_employee_id(
        record.get("employee_id")
    )

    name = str(
        record.get("name") or ""
    ).strip()

    days = float(
        record.get("days_present") or 0
    )

    return (
        f"{employee_id} | {name} | "
        f"{days:g} days"
    )


def _mapping_key_to_record(
    attendance_records: List[Dict[str, Any]],
) -> Dict[str, Dict[str, Any]]:
    result = {}

    for record in attendance_records:
        employee_id = normalize_employee_id(
            record.get("employee_id")
        )

        if employee_id:
            result[employee_id] = record

    return result


def _format_days(value: Any) -> str:
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return "0"

    if numeric.is_integer():
        return str(int(numeric))

    return f"{numeric:.2f}".rstrip("0").rstrip(".")


def _initialize_mapping_state(
    template_data,
    attendance_records,
    register,
) -> Dict[str, Dict[str, Any]]:
    resolved = resolve_template_rows(
        template_data=template_data,
        attendance_records=attendance_records,
        register=register,
    )

    return {
        mapping["mapping_key"]: mapping
        for mapping in resolved
    }


def _mapping_is_blocking(
    mapping: Dict[str, Any],
) -> bool:
    return bool(
        mapping.get("needs_confirmation")
        and not mapping.get("confirmed")
    )


def _render_identity_stats(
    mappings: List[Dict[str, Any]],
) -> None:
    total = len(mappings)

    id_matches = sum(
        mapping["status"]
        in {
            STATUS_REGISTERED,
            STATUS_REGISTERED_NAME_VARIANT,
        }
        for mapping in mappings
        if mapping.get("confirmed")
    )

    exact_matches = sum(
        mapping["status"] == "EXACT NAME MATCH"
        and mapping.get("confirmed")
        for mapping in mappings
    )

    reviews = sum(
        _mapping_is_blocking(mapping)
        for mapping in mappings
    )

    missing = sum(
        mapping["status"]
        == STATUS_MISSING_ATTENDANCE
        for mapping in mappings
    )

    not_present = sum(
        mapping["status"]
        == STATUS_NOT_PRESENT
        for mapping in mappings
    )

    conflicts = validate_mapping_conflicts(
        mappings
    )

    cols = st.columns(6)

    cols[0].metric(
        "Payroll Rows",
        total,
    )

    cols[1].metric(
        "ID Matched",
        id_matches,
    )

    cols[2].metric(
        "Exact Name",
        exact_matches,
    )

    cols[3].metric(
        "Needs Review",
        reviews,
    )

    cols[4].metric(
        "Not Present",
        not_present + missing,
    )

    cols[5].metric(
        "ID Conflicts",
        len(conflicts),
    )


def _current_selected_id(
    mapping: Dict[str, Any],
) -> str:
    return normalize_employee_id(
        mapping.get("essl_id")
    )


def _set_mapping_from_attendance(
    mapping: Dict[str, Any],
    employee_id: str,
    attendance_by_id: Dict[str, Dict[str, Any]],
    *,
    user_confirmed: bool = False,
) -> None:
    employee_id = normalize_employee_id(
        employee_id
    )

    mapping["essl_id"] = employee_id

    if employee_id and employee_id in attendance_by_id:
        record = attendance_by_id[
            employee_id
        ]

        mapping["attendance_name"] = str(
            record.get("name") or ""
        ).strip()

        mapping["days_present"] = float(
            record.get("days_present") or 0
        )

        mapping["match_method"] = (
            "user_confirmed"
            if user_confirmed
            else "selected_attendance_id"
        )

        mapping["status"] = (
            "USER CONFIRMED"
            if user_confirmed
            else STATUS_REVIEW
        )

        mapping["confirmed"] = bool(
            user_confirmed
        )

        mapping["needs_confirmation"] = (
            not user_confirmed
        )

        mapping["note"] = (
            "Selected by ESSL Employee ID."
            if user_confirmed
            else "Select 'Confirm Identity' to use this ID."
        )

    else:
        mapping["attendance_name"] = ""
        mapping["days_present"] = 0.0
        mapping["match_method"] = (
            "no_attendance_match"
        )
        mapping["status"] = STATUS_NOT_PRESENT
        mapping["confirmed"] = True
        mapping["needs_confirmation"] = False
        mapping["note"] = (
            "No attendance identity selected."
        )


def _render_mapping_row(
    mapping: Dict[str, Any],
    attendance_records: List[Dict[str, Any]],
    attendance_by_id: Dict[str, Dict[str, Any]],
) -> None:
    key = mapping["mapping_key"]

    current_id = _current_selected_id(
        mapping
    )

    # Build options once.
    option_values = [NOT_PRESENT_VALUE]
    option_labels = {
        NOT_PRESENT_VALUE:
        "Not Present (0 Days)"
    }

    for record in attendance_records:
        employee_id = normalize_employee_id(
            record.get("employee_id")
        )

        if not employee_id:
            continue

        option_values.append(
            employee_id
        )

        option_labels[
            employee_id
        ] = _mapping_option_label(
            record
        )

    if current_id and current_id not in option_values:
        current_id = ""

    selected_value = st.selectbox(
        "Attendance identity",
        options=option_values,
        index=(
            option_values.index(
                current_id
            )
            if current_id
            else 0
        ),
        format_func=lambda value: option_labels.get(
            value,
            value,
        ),
        key=f"identity_select_{key}",
        label_visibility="collapsed",
    )

    if selected_value == NOT_PRESENT_VALUE:
        if mapping.get("essl_id"):
            mapping["essl_id"] = ""
            mapping["attendance_name"] = ""
            mapping["days_present"] = 0.0
            mapping["status"] = STATUS_NOT_PRESENT
            mapping["match_method"] = (
                "no_attendance_match"
            )
            mapping["confirmed"] = True
            mapping["needs_confirmation"] = False
            mapping["note"] = (
                "No attendance record assigned."
            )

    elif selected_value != current_id:
        # New selection always requires explicit confirmation unless
        # the row is already backed by a persistent ID and the user
        # has intentionally changed it.
        _set_mapping_from_attendance(
            mapping,
            selected_value,
            attendance_by_id,
            user_confirmed=False,
        )

    status = mapping.get(
        "status",
        STATUS_NOT_PRESENT,
    )

    requires_confirmation = bool(
        mapping.get("needs_confirmation")
        and mapping.get("essl_id")
    )

    if requires_confirmation:
        confirm_key = (
            f"confirm_identity_{key}"
        )

        confirmed = st.checkbox(
            "Confirm Identity",
            value=bool(
                mapping.get("confirmed")
            ),
            key=confirm_key,
        )

        if confirmed:
            mapping["confirmed"] = True
            mapping["needs_confirmation"] = False
            mapping["status"] = (
                "USER CONFIRMED"
            )
            mapping["match_method"] = (
                "user_confirmed"
            )
            mapping["note"] = (
                "Identity explicitly confirmed by user."
            )
        else:
            mapping["confirmed"] = False
            mapping["needs_confirmation"] = True

    status = mapping.get(
        "status",
        STATUS_NOT_PRESENT,
    )

    sheet_name = mapping[
        "sheet_name"
    ]

    section_type = mapping[
        "section_type"
    ]

    row_num = mapping[
        "row_num"
    ]

    payroll_name = mapping[
        "payroll_name"
    ]

    payroll_code = mapping.get(
        "payroll_code"
    )

    attendance_name = mapping.get(
        "attendance_name"
    )

    essl_id = mapping.get(
        "essl_id"
    )

    days_present = mapping.get(
        "days_present",
        0,
    )

    cols = st.columns(
        [1.8, 2.0, 1.0, 1.8]
    )

    cols[0].markdown(
        f"**{payroll_name}**  \n"
        f"`{sheet_name}` · {section_type} · row {row_num}"
    )

    cols[1].markdown(
        f"Payroll code: `{payroll_code or 'N/A'}`  \n"
        f"ESSL ID: `{essl_id or 'None'}`"
    )

    cols[2].markdown(
        f"**{_format_days(days_present)}**  \n"
        f"days"
    )

    status_display = status

    if status_display in {
        STATUS_REVIEW,
        "MULTIPLE ID CANDIDATES",
    }:
        cols[3].warning(
            f"{status_display}"
        )
    elif status_display == STATUS_NOT_PRESENT:
        cols[3].info(
            "Not present"
        )
    else:
        cols[3].success(
            status_display
        )

    if attendance_name:
        st.caption(
            f"Attendance name: {attendance_name}"
        )

    note = mapping.get(
        "note"
    )

    if note:
        st.caption(
            note
        )

    st.markdown(
        "---"
    )


def _calculate_preview_rows(
    template_data,
    mappings_by_key,
    total_days,
    new_employees,
):
    rows = []

    for sheet_name, sections in template_data.items():

        for section in sections:

            for emp in section.get(
                "rows",
                [],
            ):

                key = payroll_mapping_key(
                    sheet_name,
                    section["type"],
                    emp["row_num"],
                )

                mapping = mappings_by_key.get(
                    key,
                    {},
                )

                if (
                    mapping.get("confirmed")
                    and mapping.get("essl_id")
                ):
                    days_present = float(
                        mapping.get(
                            "days_present",
                            0,
                        )
                        or 0
                    )
                else:
                    days_present = 0.0

                gross = float(
                    emp.get("gross_salary")
                    or 0
                )

                prorated = (
                    gross
                    * (
                        days_present
                        / total_days
                    )
                    if total_days
                    else 0
                )

                if section["type"] == "employee":

                    deductions = 200.0

                    net = (
                        prorated
                        - deductions
                    )

                    record_type = "Employee"

                else:

                    deductions = (
                        prorated
                        * 0.10
                    )

                    net = (
                        prorated
                        - deductions
                    )

                    record_type = (
                        "Freelancer (TDS)"
                    )

                rows.append(
                    {
                        "Sheet": sheet_name,
                        "Type": record_type,
                        "Name": emp["name"],
                        "ESSL ID": mapping.get(
                            "essl_id",
                            "",
                        ),
                        "Gross Base": gross,
                        "Present Days": days_present,
                        "Pro-rated Gross": round(
                            prorated,
                            2,
                        ),
                        "TDS / Prof Tax": round(
                            deductions,
                            2,
                        ),
                        "Net Salary": round(
                            net,
                            2,
                        ),
                    }
                )

    for hire in new_employees:

        days_present = float(
            hire.get(
                "days_present",
                0,
            )
            or 0
        )

        gross = float(
            hire.get(
                "gross_salary",
                0,
            )
            or 0
        )

        prorated = (
            gross
            * (
                days_present
                / total_days
            )
            if total_days
            else 0
        )

        if hire["type"] == "employee":
            deductions = 200.0
            net = prorated - deductions
        else:
            deductions = prorated * 0.10
            net = prorated - deductions

        rows.append(
            {
                "Sheet": hire["sheet"],
                "Type": (
                    f"New "
                    f"{hire['type'].capitalize()}"
                ),
                "Name": hire["name"],
                "ESSL ID": hire.get(
                    "essl_id",
                    "",
                ),
                "Gross Base": gross,
                "Present Days": days_present,
                "Pro-rated Gross": round(
                    prorated,
                    2,
                ),
                "TDS / Prof Tax": round(
                    deductions,
                    2,
                ),
                "Net Salary": round(
                    net,
                    2,
                ),
            }
        )

    return rows


def _write_existing_rows(
    wb,
    template_data,
    mappings_by_key,
    total_days,
):
    attendance_updates = []

    for sheet_name, sections in template_data.items():

        ws = wb[sheet_name]

        for section in sections:

            for emp in section.get(
                "rows",
                [],
            ):

                key = payroll_mapping_key(
                    sheet_name,
                    section["type"],
                    emp["row_num"],
                )

                mapping = mappings_by_key.get(
                    key,
                    {},
                )

                days_present = (
                    float(
                        mapping.get(
                            "days_present",
                            0,
                        )
                        or 0
                    )
                    if mapping.get("confirmed")
                    else 0.0
                )

                row_num = int(
                    emp["row_num"]
                )

                if section["type"] == "employee":

                    # E = total days
                    # F = present days
                    ws.cell(
                        row_num,
                        5,
                    ).value = total_days

                    ws.cell(
                        row_num,
                        6,
                    ).value = days_present

                    # Clear loan / advance.
                    ws.cell(
                        row_num,
                        19,
                    ).value = None

                else:

                    # D = total days
                    # E = present days
                    ws.cell(
                        row_num,
                        4,
                    ).value = total_days

                    ws.cell(
                        row_num,
                        5,
                    ).value = days_present

                    # Clear loan / advance.
                    ws.cell(
                        row_num,
                        15,
                    ).value = None

                attendance_updates.append(
                    (
                        sheet_name,
                        section,
                        emp,
                        mapping,
                    )
                )

    return attendance_updates


def _collect_register_mappings(
    template_data,
    mappings_by_key,
    new_hire_results,
    payroll_month,
):
    mappings = []

    for mapping in mappings_by_key.values():

        if not mapping.get(
            "confirmed"
        ):
            continue

        essl_id = normalize_employee_id(
            mapping.get("essl_id")
        )

        if not essl_id:
            continue

        mappings.append(
            {
                "essl_id": essl_id,
                "attendance_name": mapping.get(
                    "attendance_name",
                    "",
                ),
                "sheet_name": mapping[
                    "sheet_name"
                ],
                "section_type": mapping[
                    "section_type"
                ],
                "row_num": mapping[
                    "row_num"
                ],
                "payroll_code": mapping.get(
                    "payroll_code"
                ),
                "payroll_name": mapping[
                    "payroll_name"
                ],
                "match_method": mapping.get(
                    "match_method",
                    "user_confirmed",
                ),
                "confirmed": True,
            }
        )

    for result in new_hire_results:

        mappings.append(
            {
                "essl_id": normalize_employee_id(
                    result["essl_id"]
                ),
                "attendance_name": result[
                    "attendance_name"
                ],
                "sheet_name": result[
                    "sheet_name"
                ],
                "section_type": result[
                    "section_type"
                ],
                "row_num": result[
                    "row_num"
                ],
                "payroll_code": result.get(
                    "payroll_code"
                ),
                "payroll_name": result[
                    "payroll_name"
                ],
                "match_method": "new_hire_confirmed",
                "confirmed": True,
            }
        )

    return mappings


# ---------------------------------------------------------------------------
# SIDEBAR INPUTS
# ---------------------------------------------------------------------------

with st.sidebar:

    st.markdown(
        "### Payroll Configuration"
    )

    st.markdown(
        "**1. Upload Payroll Template**"
    )

    uploaded_template = st.file_uploader(
        "Excel template",
        type=["xlsx"],
        key="template_uploader",
    )

    st.markdown("---")

    st.markdown(
        "**2. Upload Monthly Attendance**"
    )

    uploaded_attendance = st.file_uploader(
        "Attendance workbook",
        type=["xlsx", "xls"],
        key="attendance_uploader",
    )

    st.markdown("---")

    archive_enabled = st.checkbox(
        "Archive previous month's consultants",
        value=True,
    )


# ---------------------------------------------------------------------------
# REQUIRE INPUTS
# ---------------------------------------------------------------------------

if (
    uploaded_template is None
    or uploaded_attendance is None
):

    st.info(
        "Upload both the payroll template and the monthly attendance workbook."
    )

    st.stop()


# ---------------------------------------------------------------------------
# FILE KEYS AND RESET
# ---------------------------------------------------------------------------

attendance_key = (
    uploaded_attendance.name,
    uploaded_attendance.size,
)

template_key = (
    uploaded_template.name,
    uploaded_template.size,
)


if attendance_key != st.session_state.attendance_key:

    st.session_state.attendance_key = (
        attendance_key
    )

    st.session_state.attendance_records = []

    st.session_state.mapping_state = {}

    st.session_state.new_employees = []

    st.session_state.generated_file = None

    st.session_state.download_filename = None


if template_key != st.session_state.template_key:

    st.session_state.template_key = (
        template_key
    )

    st.session_state.template_data = {}

    st.session_state.mapping_state = {}

    st.session_state.new_employees = []

    st.session_state.generated_file = None

    st.session_state.download_filename = None


# ---------------------------------------------------------------------------
# LOAD ATTENDANCE
# ---------------------------------------------------------------------------

if not st.session_state.attendance_records:

    try:

        month_from_attendance, records = (
            parse_attendance_input(
                uploaded_attendance,
                filename=uploaded_attendance.name,
            )
        )

        st.session_state.attendance_records = (
            records
        )

    except Exception as exc:

        st.error(
            f"Error reading attendance workbook: {exc}"
        )

        st.stop()


attendance_records = (
    st.session_state.attendance_records
)


# ---------------------------------------------------------------------------
# LOAD TEMPLATE
# ---------------------------------------------------------------------------

try:

    wb_preview = _load_workbook_from_source(
        uploaded_template,
        read_only=True,
    )

    template_data = _template_sections(
        wb_preview
    )

    template_month = detect_workbook_month(
        wb_preview
    )

    wb_preview.close()

except Exception as exc:

    st.error(
        f"Error loading payroll template: {exc}"
    )

    st.stop()


st.session_state.template_data = (
    template_data
)

st.session_state.template_month = (
    template_month
)


# ---------------------------------------------------------------------------
# TARGET MONTH
# ---------------------------------------------------------------------------

default_month = (
    st.session_state.target_month
    or template_month
)

if not default_month:

    month_from_filename = (
        month_from_attendance
        if "month_from_attendance" in locals()
        else None
    )

    default_month = (
        month_from_filename
        or "August 2026"
    )


st.session_state.target_month = (
    st.text_input(
        "Target Payroll Month & Year",
        value=default_month,
        key="target_month_input",
    )
)


total_days = get_days_in_month(
    st.session_state.target_month
)


# ---------------------------------------------------------------------------
# IDENTITY REGISTER
# ---------------------------------------------------------------------------

try:

    wb_identity_preview = _load_workbook_from_source(
        uploaded_template,
        read_only=False,
    )

    identity_register = load_identity_register(
        wb_identity_preview
    )

    wb_identity_preview.close()

except Exception as exc:

    st.error(
        f"Error loading identity register: {exc}"
    )

    st.stop()


st.session_state.identity_register_count = (
    len(identity_register)
)


# ---------------------------------------------------------------------------
# INITIALIZE IDENTITY MATCHING
# ---------------------------------------------------------------------------

if not st.session_state.mapping_state:

    st.session_state.mapping_state = (
        _initialize_mapping_state(
            template_data=template_data,
            attendance_records=attendance_records,
            register=identity_register,
        )
    )


mappings_by_key = (
    st.session_state.mapping_state
)

mapping_list = list(
    mappings_by_key.values()
)


# ---------------------------------------------------------------------------
# TABS
# ---------------------------------------------------------------------------

tab_identity, tab_preview, tab_generate = (
    st.tabs(
        [
            "Identity & Attendance",
            "Payroll Preview",
            "Generate & Download",
        ]
    )
)


# ---------------------------------------------------------------------------
# TAB 1: IDENTITY & ATTENDANCE
# ---------------------------------------------------------------------------

with tab_identity:

    st.subheader(
        "Employee Identity & Attendance"
    )

    if identity_register:

        st.success(
            f"Persistent identity register loaded: "
            f"{len(identity_register)} mapping(s)."
        )

    else:

        st.info(
            "This template does not yet contain an identity register. "
            "Exact unique name matches are used only to initialize "
            "the register. Name similarity suggestions require "
            "explicit confirmation."
        )

    _render_identity_stats(
        mapping_list
    )

    conflicts = validate_mapping_conflicts(
        mapping_list
    )

    if conflicts:

        st.error(
            "The same ESSL Employee ID is assigned to multiple "
            "payroll rows. Generation is blocked until the conflict "
            "is resolved."
        )

    review_count = sum(
        _mapping_is_blocking(mapping)
        for mapping in mapping_list
    )

    if review_count:

        st.warning(
            f"{review_count} payroll row(s) require identity confirmation "
            "before payroll can be generated."
        )

    st.markdown(
        "### Identity Mapping"
    )

    attendance_by_id, _ = (
        build_attendance_indexes(
            attendance_records
        )
    )

    for sheet_name, sections in (
        template_data.items()
    ):

        st.markdown(
            f"#### {sheet_name}"
        )

        for section in sections:

            st.markdown(
                f"**{section['type'].capitalize()}**"
            )

            for emp in section.get(
                "rows",
                [],
            ):

                key = payroll_mapping_key(
                    sheet_name,
                    section["type"],
                    emp["row_num"],
                )

                mapping = mappings_by_key.get(
                    key
                )

                if mapping is None:
                    continue

                _render_mapping_row(
                    mapping,
                    attendance_records,
                    attendance_by_id,
                )

    # ---------------------------------------------------------
    # UNASSIGNED ATTENDANCE
    # ---------------------------------------------------------

    confirmed_ids = {
        normalize_employee_id(
            mapping.get("essl_id")
        )
        for mapping in mapping_list
        if mapping.get("confirmed")
        and mapping.get("essl_id")
    }

    assigned_new_hire_ids = {
        normalize_employee_id(
            hire.get("essl_id")
        )
        for hire in st.session_state.new_employees
        if hire.get("essl_id")
    }

    used_ids = (
        confirmed_ids
        | assigned_new_hire_ids
    )

    unassigned_records = [
        record
        for record in attendance_records
        if normalize_employee_id(
            record.get("employee_id")
        )
        not in used_ids
    ]

    st.markdown(
        "### Unassigned Attendance Records"
    )

    st.caption(
        "These ESSL records were not assigned to an existing payroll row."
    )

    if unassigned_records:

        unassigned_df = pd.DataFrame(
            [
                {
                    "ESSL ID": normalize_employee_id(
                        record.get("employee_id")
                    ),
                    "Attendance Name": record.get(
                        "name"
                    ),
                    "Days Present": record.get(
                        "days_present"
                    ),
                }
                for record in unassigned_records
            ]
        )

        st.dataframe(
            unassigned_df,
            use_container_width=True,
            hide_index=True,
        )

    else:

        st.success(
            "All attendance identities are currently assigned or intentionally excluded."
        )

    # ---------------------------------------------------------
    # NEW HIRE FORM
    # ---------------------------------------------------------

    st.markdown(
        "### Add New Employee / Freelancer"
    )

    with st.form(
        "new_hire_form"
    ):

        hire_options = [
            NOT_PRESENT_VALUE
        ]

        hire_labels = {
            NOT_PRESENT_VALUE:
            "Select attendance record"
        }

        for record in unassigned_records:

            employee_id = normalize_employee_id(
                record.get("employee_id")
            )

            if not employee_id:
                continue

            hire_options.append(
                employee_id
            )

            hire_labels[
                employee_id
            ] = _mapping_option_label(
                record
            )

        col1, col2, col3 = st.columns(3)

        with col1:

            hire_id = st.selectbox(
                "Attendance identity",
                hire_options,
                format_func=lambda value: (
                    hire_labels.get(
                        value,
                        value,
                    )
                ),
            )

        with col2:

            target_sheet = st.selectbox(
                "Payroll Sheet",
                list(
                    template_data.keys()
                ),
            )

        with col3:

            hire_type = st.selectbox(
                "Type",
                [
                    "Employee",
                    "Freelancer (TDS)",
                ],
            )

        col4, col5, col6 = st.columns(3)

        with col4:

            designation = st.text_input(
                "Designation"
            )

        with col5:

            branch = st.text_input(
                "Branch / Function"
            )

        with col6:

            gross_salary = st.number_input(
                "Gross Salary",
                min_value=0.0,
                step=100.0,
            )

        submitted = st.form_submit_button(
            "Add New Hire"
        )

        if submitted:

            if hire_id == NOT_PRESENT_VALUE:

                st.error(
                    "Select an attendance identity."
                )

            elif not designation.strip():

                st.error(
                    "Designation is required."
                )

            elif not branch.strip():

                st.error(
                    "Branch / Function is required."
                )

            else:

                record = attendance_by_id[
                    hire_id
                ]

                selected_type = (
                    "employee"
                    if hire_type
                    == "Employee"
                    else "consultant"
                )

                hire = {
                    "essl_id": hire_id,
                    "attendance_name": record[
                        "name"
                    ],
                    "days_present": float(
                        record.get(
                            "days_present",
                            0,
                        )
                        or 0
                    ),
                    "sheet": target_sheet,
                    "type": selected_type,
                    "designation": designation.strip(),
                    "branch": branch.strip(),
                    "gross_salary": float(
                        gross_salary
                    ),
                }

                st.session_state.new_employees.append(
                    hire
                )

                st.success(
                    f"Added {record['name']} as a new hire."
                )

                st.rerun()

    if st.session_state.new_employees:

        st.markdown(
            "#### New Hires Pending"
        )

        for index, hire in enumerate(
            st.session_state.new_employees,
            start=1,
        ):

            st.write(
                f"{index}. "
                f"{hire['attendance_name']} "
                f"(ESSL ID {hire['essl_id']}) "
                f"→ {hire['sheet']} "
                f"→ {hire['type']} "
                f"→ {hire['days_present']} days"
            )

        if st.button(
            "Clear New Hires"
        ):

            st.session_state.new_employees = []

            st.rerun()


# ---------------------------------------------------------------------------
# TAB 2: PAYROLL PREVIEW
# ---------------------------------------------------------------------------

with tab_preview:

    st.subheader(
        "Payroll Calculations Preview"
    )

    st.caption(
        "Only confirmed ESSL identities contribute attendance to payroll."
    )

    preview_rows = _calculate_preview_rows(
        template_data=template_data,
        mappings_by_key=mappings_by_key,
        total_days=total_days,
        new_employees=st.session_state.new_employees,
    )

    if preview_rows:

        df_preview = pd.DataFrame(
            preview_rows
        )

        st.dataframe(
            df_preview,
            use_container_width=True,
            hide_index=True,
        )

    else:

        st.info(
            "No payroll rows available."
        )


# ---------------------------------------------------------------------------
# TAB 3: GENERATE
# ---------------------------------------------------------------------------

with tab_generate:

    st.subheader(
        "Generate Monthly Payroll"
    )

    target_month = (
        st.session_state.target_month
    )

    current_template_month = (
        st.session_state.template_month
    )

    st.write(
        f"Target month: **{target_month}**"
    )

    if current_template_month:

        st.caption(
            f"Template currently appears to be for "
            f"**{current_template_month}**."
        )

    blocking_reviews = [
        mapping
        for mapping in mapping_list
        if _mapping_is_blocking(mapping)
    ]

    conflicts = validate_mapping_conflicts(
        mapping_list
    )

    if blocking_reviews:

        st.error(
            f"Resolve {len(blocking_reviews)} identity review(s) "
            "before generating payroll."
        )

    if conflicts:

        st.error(
            "Resolve duplicate ESSL Employee ID assignments before generating payroll."
        )

    should_archive = bool(
        archive_enabled
        and current_template_month
        and (
            current_template_month
            != target_month
        )
    )

    if archive_enabled and not should_archive:

        st.info(
            "Consultant archiving is disabled for this run because "
            "the uploaded template is already for the target month."
        )

    st.warning(
        "The payroll workbook will be modified in memory and a new "
        "downloadable workbook will be created. The uploaded template "
        "file itself is never overwritten."
    )

    can_generate = (
        not blocking_reviews
        and not conflicts
    )

    if st.button(
        "Generate Payroll Excel",
        type="primary",
        disabled=not can_generate,
    ):

        with st.spinner(
            "Generating payroll workbook..."
        ):

            try:

                wb = _load_workbook_from_source(
                    uploaded_template,
                    read_only=False,
                )

                # ---------------------------------------------------------
                # 1. Archive previous month only when moving to a new
                #    month.
                # ---------------------------------------------------------

                if should_archive:

                    archive_previous_month_consultants(
                        wb,
                        current_template_month,
                    )

                # ---------------------------------------------------------
                # 2. Update active month headers.
                #    AI - TDS is intentionally excluded because it is
                #    historical data.
                # ---------------------------------------------------------

                update_active_workbook_month(
                    wb,
                    target_month,
                    skip_sheets={
                        "AI - TDS",
                        IDENTITY_SHEET_NAME,
                    },
                )

                # ---------------------------------------------------------
                # 3. Write attendance to existing rows.
                # ---------------------------------------------------------

                _write_existing_rows(
                    wb=wb,
                    template_data=template_data,
                    mappings_by_key=mappings_by_key,
                    total_days=total_days,
                )

                # ---------------------------------------------------------
                # 4. Add new hires.
                # ---------------------------------------------------------

                new_hire_results = []

                for hire in st.session_state.new_employees:

                    sections = template_data.get(
                        hire["sheet"],
                        [],
                    )

                    target_section = next(
                        (
                            section
                            for section in sections
                            if section["type"]
                            == hire["type"]
                        ),
                        None,
                    )

                    if target_section is None:

                        raise ValueError(
                            f"No {hire['type']} section exists "
                            f"in sheet '{hire['sheet']}'."
                        )

                    inserted_row = add_new_hire_to_section(
                        wb[hire["sheet"]],
                        target_section,
                        {
                            "name": hire[
                                "attendance_name"
                            ],
                            "designation": hire[
                                "designation"
                            ],
                            "branch": hire[
                                "branch"
                            ],
                            "gross_salary": hire[
                                "gross_salary"
                            ],
                        },
                        total_days=total_days,
                        days_present=float(
                            hire[
                                "days_present"
                            ]
                        ),
                    )

                    if inserted_row is None:

                        raise ValueError(
                            f"No blank row available in "
                            f"'{hire['sheet']}' for "
                            f"{hire['attendance_name']}."
                        )

                    inserted_code = wb[
                        hire["sheet"]
                    ].cell(
                        inserted_row,
                        1,
                    ).value

                    new_hire_results.append(
                        {
                            "essl_id": hire[
                                "essl_id"
                            ],
                            "attendance_name": hire[
                                "attendance_name"
                            ],
                            "sheet_name": hire[
                                "sheet"
                            ],
                            "section_type": hire[
                                "type"
                            ],
                            "row_num": inserted_row,
                            "payroll_code": inserted_code,
                            "payroll_name": hire[
                                "attendance_name"
                            ],
                        }
                    )

                # ---------------------------------------------------------
                # 5. Persist identity register.
                # ---------------------------------------------------------

                register_mappings = (
                    _collect_register_mappings(
                        template_data=template_data,
                        mappings_by_key=mappings_by_key,
                        new_hire_results=new_hire_results,
                        payroll_month=target_month,
                    )
                )

                ensure_identity_sheet(
                    wb
                )

                upsert_identity_register(
                    wb,
                    register_mappings,
                    payroll_month=target_month,
                )

                # ---------------------------------------------------------
                # 6. Ask Excel to recalculate formulas.
                # ---------------------------------------------------------

                try:
                    wb.calculation.fullCalcOnLoad = True
                    wb.calculation.forceFullCalc = True
                    wb.calculation.calcMode = "auto"
                except Exception:
                    pass

                # ---------------------------------------------------------
                # 7. Save.
                # ---------------------------------------------------------

                output = io.BytesIO()

                wb.save(
                    output
                )

                wb.close()

                output.seek(0)

                st.session_state.generated_file = (
                    output.getvalue()
                )

                safe_month = re.sub(
                    r"[^A-Za-z0-9]+",
                    "_",
                    target_month,
                ).strip("_")

                st.session_state.download_filename = (
                    f"Payroll_{safe_month}.xlsx"
                )

                st.success(
                    "Payroll workbook generated successfully."
                )

            except Exception as exc:

                st.error(
                    f"Error generating payroll workbook: {exc}"
                )


# ---------------------------------------------------------------------------
# DOWNLOAD
# ---------------------------------------------------------------------------

if (
    st.session_state.generated_file
    is not None
):

    st.markdown(
        "### Download"
    )

    st.download_button(
        "Download Generated Payroll Excel",
        data=st.session_state.generated_file,
        file_name=(
            st.session_state.download_filename
            or "Payroll_Output.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="final_download",
    )
