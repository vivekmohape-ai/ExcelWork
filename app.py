from __future__ import annotations

import re
import streamlit as st
import pandas as pd

from excel_utils import (
    analyze_template_sheet,
    apply_daily_punch_corrections,
    apply_leave_adjustments,
    get_days_in_month,
    parse_attendance_input,
    parse_daily_punch_report,
    parse_leave_report,
)
from identity import (
    STATUS_MISSING_ATTENDANCE,
    STATUS_NOT_PRESENT,
    STATUS_REGISTERED,
    STATUS_REGISTERED_NAME_VARIANT,
    STATUS_REVIEW,
    build_attendance_indexes,
    load_identity_register,
    normalize_employee_id,
    payroll_mapping_key,
    resolve_template_rows,
    validate_mapping_conflicts,
)
from payroll_engine import (
    detect_active_sheet_names,
    detect_template_month_safe,
    generate_payroll_workbook,
    load_workbook,
    sha256_bytes,
    validate_attendance_records,
)

st.set_page_config(
    page_title="Payroll Automator",
    page_icon="💸",
    layout="wide",
)

st.title("Payroll Automator")
st.caption(
    "Payroll template + ESSL attendance + Leave / Comp Off report. "
    "Final payable days are reconciled before salary calculation."
)

DEFAULTS = {
    "template_hash": None,
    "attendance_hash": None,
    "daily_punch_hash": None,
    "leave_hash": None,
    "target_month": None,
    "attendance_records": [],
    "leave_records": [],
    "leave_summary": {},
    "template_data": {},
    "template_month": None,
    "active_sheets": [],
    "mapping_state": {},
    "new_hires": [],
    "generated_file": None,
    "generated_filename": None,
}

for key, value in DEFAULTS.items():
    st.session_state.setdefault(key, value)


def reset_for_inputs(
    template_hash: str,
    attendance_hash: str,
    daily_punch_hash: str,
    leave_hash: str,
):
    if (
        template_hash != st.session_state.template_hash
        or attendance_hash != st.session_state.attendance_hash
        or daily_punch_hash != st.session_state.daily_punch_hash
        or leave_hash != st.session_state.leave_hash
    ):
        st.session_state.template_hash = template_hash
        st.session_state.attendance_hash = attendance_hash
        st.session_state.daily_punch_hash = daily_punch_hash
        st.session_state.leave_hash = leave_hash
        st.session_state.mapping_state = {}
        st.session_state.new_hires = []
        st.session_state.generated_file = None
        st.session_state.generated_filename = None
        st.session_state.leave_summary = {}


def format_days(value) -> str:
    try:
        value = float(value or 0)
    except Exception:
        return "0"
    return str(int(value)) if value.is_integer() else f"{value:.2f}"


with st.sidebar:
    st.markdown("### Payroll Configuration")

    template_file = st.file_uploader(
        "1. Payroll template",
        type=["xlsx"],
        key="template_file",
    )

    attendance_file = st.file_uploader(
        "2. Attendance output from ESSL Extractor",
        type=["xlsx", "xls"],
        key="attendance_file",
        help=(
            "Use the Excel produced by the ESSL Extractor. "
            "It should contain Payroll Present Days."
        ),
    )

    daily_punch_file = st.file_uploader(
        "3. Daily Attendance Report for punch verification",
        type=["xlsx", "xls"],
        key="daily_punch_file",
        help=(
            "Optional ESSL Daily Attendance Report. "
            "If either InTime or OutTime exists for an employee/date, "
            "that date is corrected to Present."
        ),
    )

    leave_file = st.file_uploader(
        "4. Leave / Comp Off report",
        type=["xlsx", "xls"],
        key="leave_file",
        help=(
            "Monthly approved leave, approved comp off and unpaid leave report."
        ),
    )

    archive_enabled = st.checkbox(
        "Archive previous month's consultants",
        value=True,
    )


if (
    template_file is None
    or attendance_file is None
    or leave_file is None
):
    st.info(
        "Upload the payroll template, attendance Excel and Leave / Comp Off report."
    )
    st.stop()


template_hash = sha256_bytes(template_file)
attendance_hash = sha256_bytes(attendance_file)
daily_punch_hash = (
    sha256_bytes(daily_punch_file)
    if daily_punch_file is not None
    else ""
)
leave_hash = sha256_bytes(leave_file)

reset_for_inputs(
    template_hash,
    attendance_hash,
    daily_punch_hash,
    leave_hash,
)

# ------------------------------------------------------------
# Attendance
# ------------------------------------------------------------

try:
    attendance_month, attendance_records = parse_attendance_input(
        attendance_file,
        filename=attendance_file.name,
    )
except Exception as exc:
    st.error(f"Could not read attendance workbook: {exc}")
    st.stop()

# ------------------------------------------------------------
# Daily Attendance / Punch correction
# ------------------------------------------------------------

punch_summary = {
    "employees_corrected": 0,
    "days_corrected": 0,
    "already_present_or_paid": 0,
    "unmatched_employee_ids": 0,
    "details": [],
}

if daily_punch_file is not None:
    try:
        punch_month, punch_records = parse_daily_punch_report(
            daily_punch_file,
            filename=daily_punch_file.name,
        )

        if (
            attendance_month
            and punch_month
            and attendance_month.lower() != punch_month.lower()
        ):
            st.error(
                "Attendance and Daily Attendance Report are for different months: "
                f"{attendance_month} vs {punch_month}."
            )
            st.stop()

        punch_summary = apply_daily_punch_corrections(
            attendance_records,
            punch_records,
        )

    except Exception as exc:
        st.error(
            f"Could not read Daily Attendance Report: {exc}"
        )
        st.stop()

attendance_issues = validate_attendance_records(
    attendance_records
)

if attendance_issues:
    st.error("Attendance validation failed:")
    for issue in attendance_issues[:20]:
        st.write(f"• {issue}")
    st.stop()

st.session_state.attendance_records = attendance_records

# ------------------------------------------------------------
# Leave / Comp Off
# ------------------------------------------------------------

try:
    leave_month, leave_records = parse_leave_report(
        leave_file,
        filename=leave_file.name,
    )
except Exception as exc:
    st.error(f"Could not read Leave / Comp Off report: {exc}")
    st.stop()

st.session_state.leave_records = leave_records

if (
    attendance_month
    and leave_month
    and attendance_month.lower() != leave_month.lower()
):
    st.error(
        "Attendance and Leave / Comp Off reports are for different months: "
        f"{attendance_month} vs {leave_month}."
    )
    st.stop()

# ------------------------------------------------------------
# Payroll template
# ------------------------------------------------------------

try:
    template_wb = load_workbook(template_file)

    template_month = detect_template_month_safe(
        template_wb
    )

    active_sheets = detect_active_sheet_names(
        template_wb,
        template_month,
    )

    template_data = {
        sheet_name: analyze_template_sheet(
            template_wb[sheet_name]
        )
        for sheet_name in active_sheets
    }

    identity_register = load_identity_register(
        template_wb
    )

    template_wb.close()

except Exception as exc:
    st.error(f"Could not inspect payroll template: {exc}")
    st.stop()

st.session_state.template_month = template_month
st.session_state.active_sheets = active_sheets
st.session_state.template_data = template_data

default_month = (
    attendance_month
    or leave_month
    or st.session_state.target_month
    or template_month
    or ""
)

target_month = st.text_input(
    "Target Payroll Month & Year",
    value=default_month,
)

if not target_month:
    st.error("Target Payroll Month & Year is required.")
    st.stop()

st.session_state.target_month = target_month

total_days = get_days_in_month(
    target_month
)

# ------------------------------------------------------------
# Identity mapping
# ------------------------------------------------------------

if not st.session_state.mapping_state:
    mappings_list = resolve_template_rows(
        template_data,
        attendance_records,
        identity_register,
    )

    st.session_state.mapping_state = {
        mapping["mapping_key"]: mapping
        for mapping in mappings_list
    }

mappings = st.session_state.mapping_state

# Initial leave reconciliation so the Identity tab can display
# the corrected payroll days.
leave_summary = apply_leave_adjustments(
    mappings,
    leave_records,
    total_days,
)
st.session_state.leave_summary = leave_summary

# ------------------------------------------------------------
# Run summary
# ------------------------------------------------------------

c1, c2, c3, c4, c5 = st.columns(5)

c1.metric(
    "Attendance Records",
    len(attendance_records),
)

c2.metric(
    "Punch Corrections",
    int(punch_summary["days_corrected"]),
)

c3.metric(
    "Leave Records",
    len(leave_records),
)

c4.metric(
    "Active Payroll Sheets",
    len(active_sheets),
)

c5.metric(
    "Days in Month",
    total_days,
)

if daily_punch_file is not None:
    st.caption(
        f"Daily punch verification: "
        f"{punch_summary['employees_corrected']} employee(s) corrected | "
        f"{punch_summary['days_corrected']} day(s) changed to Present | "
        f"{punch_summary['already_present_or_paid']} already Present/paid dates not double-counted."
    )
else:
    st.caption(
        "Daily punch verification report not uploaded. "
        "ESSL attendance is being used without punch correction."
    )

st.caption(
    f"Leave matched: {leave_summary['matched']} | "
    f"Leave adjusted: {leave_summary['adjusted']} | "
    f"Leave unmatched/review: {leave_summary['unmatched']}"
)

if punch_summary["details"]:
    with st.expander("View Daily Punch Corrections"):
        correction_rows = []
        for item in punch_summary["details"]:
            for correction in item["corrections"]:
                correction_rows.append(
                    {
                        "ESSL ID": item["employee_id"],
                        "Employee": item["name"],
                        "Date": correction["date"],
                        "InTime": correction["in_time"],
                        "OutTime": correction["out_time"],
                        "Reason": correction["reason"],
                    }
                )

        if correction_rows:
            st.dataframe(
                pd.DataFrame(correction_rows),
                use_container_width=True,
                hide_index=True,
            )

attendance_by_id, _ = build_attendance_indexes(
    attendance_records
)

attendance_options = [
    normalize_employee_id(
        record.get("employee_id")
    )
    for record in attendance_records
    if normalize_employee_id(
        record.get("employee_id")
    )
]

attendance_labels = {
    normalize_employee_id(record.get("employee_id")):
        (
            f"{normalize_employee_id(record.get('employee_id'))} | "
            f"{record.get('name')} | "
            f"{format_days(record.get('payroll_present_days', record.get('days_present')))} "
            "payroll days"
        )
    for record in attendance_records
    if normalize_employee_id(record.get("employee_id"))
}

conflicts = validate_mapping_conflicts(
    list(mappings.values())
)

review_count = sum(
    1
    for mapping in mappings.values()
    if mapping.get("needs_confirmation")
    and not mapping.get("confirmed")
)

id_match_count = sum(
    1
    for mapping in mappings.values()
    if mapping.get("confirmed")
    and mapping.get("essl_id")
    and mapping.get("status") in {
        STATUS_REGISTERED,
        STATUS_REGISTERED_NAME_VARIANT,
        "USER CONFIRMED",
        "EXACT NAME MATCH",
    }
)

no_attendance_count = sum(
    1
    for mapping in mappings.values()
    if mapping.get("status") in {
        STATUS_NOT_PRESENT,
        STATUS_MISSING_ATTENDANCE,
    }
)

mc1, mc2, mc3, mc4 = st.columns(4)

mc1.metric(
    "ID / Exact Matched",
    id_match_count,
)

mc2.metric(
    "Needs Review",
    review_count,
)

mc3.metric(
    "No Attendance",
    no_attendance_count,
)

mc4.metric(
    "ID Conflicts",
    len(conflicts),
)

if conflicts:
    st.error(
        "The same ESSL Employee ID is assigned to more than one payroll row."
    )

# ------------------------------------------------------------
# Tabs
# ------------------------------------------------------------

tab1, tab2, tab3 = st.tabs(
    [
        "Identity & Attendance",
        "Payroll Preview",
        "Generate",
    ]
)

with tab1:
    st.subheader("Identity Mapping")

    if identity_register:
        st.success(
            f"Persistent identity register found: {len(identity_register)} mapping(s)."
        )
    else:
        st.info(
            "First run: exact normalized names are matched automatically. "
            "Name variants require confirmation."
        )

    for sheet_name, sections in template_data.items():
        st.markdown(f"#### {sheet_name}")

        for section in sections:
            st.markdown(
                f"**{section['type'].capitalize()}**"
            )

            for emp in section.get("rows", []):
                key = payroll_mapping_key(
                    sheet_name,
                    section["type"],
                    int(emp["row_num"]),
                )

                mapping = mappings[key]
                status = mapping.get("status")
                current_id = normalize_employee_id(
                    mapping.get("essl_id")
                )

                cols = st.columns(
                    [2.3, 2.5, 1.3, 1.6]
                )

                cols[0].markdown(
                    f"**{mapping.get('payroll_name')}**\n\n"
                    f"{sheet_name} · row {emp['row_num']}"
                )

                cols[1].markdown(
                    f"Payroll code: `{mapping.get('payroll_code') or 'N/A'}`\n\n"
                    f"ESSL ID: `{current_id or 'None'}`"
                )

                cols[2].markdown(
                    f"**{format_days(mapping.get('final_payable_days'))}**\n\n"
                    "final payable days"
                )

                if mapping.get("needs_confirmation"):
                    cols[3].warning(
                        status or "REVIEW"
                    )

                    values = (
                        ["__NONE__"]
                        + attendance_options
                    )

                    labels = {
                        "__NONE__":
                            "Select attendance identity"
                    }

                    labels.update(
                        attendance_labels
                    )

                    selected = st.selectbox(
                        "Attendance identity",
                        values,
                        index=(
                            values.index(current_id)
                            if current_id in values
                            else 0
                        ),
                        format_func=lambda value:
                            labels.get(value, value),
                        key=f"map_select_{key}",
                    )

                    if selected != "__NONE__":
                        if selected != current_id:
                            record = attendance_by_id[selected]

                            mapping["essl_id"] = selected
                            mapping["attendance_name"] = str(
                                record.get("name") or ""
                            )
                            mapping["payroll_present_days"] = float(
                                record.get(
                                    "payroll_present_days",
                                    record.get(
                                        "days_present",
                                        0,
                                    ),
                                )
                                or 0
                            )
                            mapping["days_present"] = (
                                mapping["payroll_present_days"]
                            )
                            mapping["confirmed"] = False
                            mapping["needs_confirmation"] = True
                            mapping["status"] = STATUS_REVIEW
                            mapping["match_method"] = "user_selected"
                            mapping["note"] = (
                                "Attendance identity selected by user. "
                                "Confirmation required."
                            )
                            st.rerun()

                    confirmed = st.checkbox(
                        "Confirm identity",
                        value=bool(
                            mapping.get("confirmed")
                        ),
                        key=f"confirm_{key}",
                    )

                    mapping["confirmed"] = confirmed
                    mapping["needs_confirmation"] = not confirmed

                    if confirmed:
                        mapping["status"] = "USER CONFIRMED"
                        mapping["match_method"] = "user_confirmed"
                        mapping["note"] = (
                            "Identity explicitly confirmed by user."
                        )

                else:
                    if status in {
                        STATUS_REGISTERED,
                        STATUS_REGISTERED_NAME_VARIANT,
                        "EXACT NAME MATCH",
                        "USER CONFIRMED",
                    }:
                        cols[3].success(status)
                    elif status == STATUS_MISSING_ATTENDANCE:
                        cols[3].info("Missing this month")
                    else:
                        cols[3].info(
                            status or "Not present"
                        )

                if mapping.get("attendance_name"):
                    st.caption(
                        f"Attendance: {mapping['attendance_name']} | "
                        f"ESSL ID: {mapping.get('essl_id')} | "
                        f"ESSL Payroll Days: "
                        f"{format_days(mapping.get('payroll_present_days'))} | "
                        f"Final Payable Days: "
                        f"{format_days(mapping.get('final_payable_days'))}"
                    )

                if (
                    mapping.get("approved_leaves")
                    or mapping.get("approved_comp_offs")
                    or mapping.get("unpaid_leaves")
                ):
                    st.caption(
                        f"Approved Leave: "
                        f"{format_days(mapping.get('approved_leaves'))} | "
                        f"Comp Off: "
                        f"{format_days(mapping.get('approved_comp_offs'))} | "
                        f"Unpaid Leave: "
                        f"{format_days(mapping.get('unpaid_leaves'))} | "
                        f"Adjustment Applied: "
                        f"{format_days(mapping.get('leave_adjustment'))}"
                    )

                if mapping.get("leave_note"):
                    st.caption(
                        f"Leave reconciliation: "
                        f"{mapping['leave_note']}"
                    )

                if mapping.get("note"):
                    st.caption(
                        mapping["note"]
                    )

                st.markdown("---")


# Recalculate after any user mapping changes.
leave_summary = apply_leave_adjustments(
    mappings,
    leave_records,
    total_days,
)
st.session_state.leave_summary = leave_summary

with tab2:
    st.subheader("Payroll Preview")

    preview_rows = []

    for mapping in mappings.values():
        confirmed = bool(
            mapping.get("confirmed")
            and mapping.get("essl_id")
        )

        days = (
            float(
                mapping.get("final_payable_days")
                or 0
            )
            if confirmed
            else 0.0
        )

        gross = float(
            mapping.get("gross_salary")
            or 0
        )

        prorated = (
            gross * days / total_days
            if total_days
            else 0
        )

        if mapping.get(
            "section_type"
        ) == "employee":
            deduction = 200.0
            kind = "Employee"
        else:
            deduction = prorated * 0.10
            kind = "Freelancer (TDS)"

        preview_rows.append(
            {
                "Sheet": mapping.get("sheet_name"),
                "Type": kind,
                "Name": mapping.get("payroll_name"),
                "ESSL ID": mapping.get("essl_id", ""),
                "ESSL Payroll Days": round(
                    float(
                        mapping.get(
                            "payroll_present_days"
                        )
                        or 0
                    ),
                    2,
                ),
                "Approved Leave": round(
                    float(
                        mapping.get(
                            "approved_leaves"
                        )
                        or 0
                    ),
                    2,
                ),
                "Comp Off": round(
                    float(
                        mapping.get(
                            "approved_comp_offs"
                        )
                        or 0
                    ),
                    2,
                ),
                "Unpaid Leave": round(
                    float(
                        mapping.get(
                            "unpaid_leaves"
                        )
                        or 0
                    ),
                    2,
                ),
                "Leave Adjustment": round(
                    float(
                        mapping.get(
                            "leave_adjustment"
                        )
                        or 0
                    ),
                    2,
                ),
                "Final Payable Days": round(
                    days,
                    2,
                ),
                "Gross Base": gross,
                "Pro-rated Gross": round(
                    prorated,
                    2,
                ),
                "TDS / Prof Tax": round(
                    deduction,
                    2,
                ),
                "Net Salary": round(
                    prorated - deduction,
                    2,
                ),
            }
        )

    if preview_rows:
        st.dataframe(
            pd.DataFrame(preview_rows),
            use_container_width=True,
            hide_index=True,
        )

with tab3:
    st.subheader("Generate Monthly Payroll")

    unresolved = [
        mapping
        for mapping in mappings.values()
        if mapping.get("needs_confirmation")
        and not mapping.get("confirmed")
    ]

    if unresolved:
        st.warning(
            f"{len(unresolved)} identity review(s) still require confirmation."
        )

    if conflicts:
        st.error(
            "Generation is blocked because of duplicate confirmed ESSL IDs."
        )

    can_generate = not conflicts

    if st.button(
        "Generate Payroll Excel",
        type="primary",
        disabled=not can_generate,
    ):
        with st.spinner(
            "Generating payroll workbook..."
        ):
            try:
                # Daily punch corrections have already been applied to
                # attendance_records before identity mapping. Make absolutely
                # sure the values used by payroll are the latest leave-reconciled values.
                apply_leave_adjustments(
                    mappings,
                    leave_records,
                    total_days,
                )

                generated, audit_rows, active = (
                    generate_payroll_workbook(
                        template_source=template_file,
                        attendance_source=attendance_file,
                        target_month=st.session_state.target_month,
                        mappings_by_key=mappings,
                        new_hires=st.session_state.new_hires,
                        archive_enabled=archive_enabled,
                    )
                )

                safe_month = re.sub(
                    r"[^A-Za-z0-9]+",
                    "_",
                    st.session_state.target_month,
                ).strip("_")

                st.session_state.generated_file = (
                    generated
                )

                st.session_state.generated_filename = (
                    f"Payroll_{safe_month}.xlsx"
                )

                st.success(
                    "Payroll generated successfully for "
                    f"{st.session_state.target_month}."
                )

                st.info(
                    f"Processed active sheets: {', '.join(active)}"
                )

            except Exception as exc:
                st.error(
                    f"Payroll generation failed: {exc}"
                )


if st.session_state.generated_file:
    st.download_button(
        "Download Generated Payroll Excel",
        data=st.session_state.generated_file,
        file_name=(
            st.session_state.generated_filename
            or "Payroll_Output.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )
