from __future__ import annotations

import hashlib
import io
import re
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from excel_utils import analyze_template_sheet, get_days_in_month, infer_month_year, parse_attendance_input
from identity import (
    IDENTITY_SHEET_NAME,
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
    initial_sidebar_state="expanded",
)

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
.review-card {
    padding: 0.9rem 1rem;
    border-radius: 10px;
    border: 1px solid #f59e0b;
    background: #fffbeb;
    margin-bottom: 0.75rem;
}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown("<h1 class='header-gradient'>Payroll Automator</h1>", unsafe_allow_html=True)
st.markdown(
    "<p class='subheader-text'>Monthly payroll processing using persistent ESSL employee identity.</p>",
    unsafe_allow_html=True,
)

DEFAULTS = {
    "template_hash": None,
    "attendance_hash": None,
    "punch_hash": None,
    "target_month": None,
    "attendance_records": [],
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


def reset_for_inputs(template_hash: str, attendance_hash: str, punch_hash: str):
    if (
        template_hash != st.session_state.template_hash
        or attendance_hash != st.session_state.attendance_hash
        or punch_hash != st.session_state.punch_hash
    ):
        st.session_state.template_hash = template_hash
        st.session_state.attendance_hash = attendance_hash
        st.session_state.punch_hash = punch_hash
        st.session_state.mapping_state = {}
        st.session_state.new_hires = []
        st.session_state.generated_file = None
        st.session_state.generated_filename = None


def format_days(value: Any) -> str:
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
        "2. Monthly attendance summary / daily attendance",
        type=["xlsx", "xls"],
        key="attendance_file",
        help="Use the normal attendance summary or a daily ESSL attendance export. If you want missed-punch correction, also upload the raw punch log below.",
    )
    punch_file = st.file_uploader(
        "3. Raw punch log (optional, recommended)",
        type=["xlsx", "xls", "csv"],
        key="punch_file",
        help="Any day with either Punch In OR Punch Out will count as a present day, even when the summary marks the day absent.",
    )
    archive_enabled = st.checkbox(
        "Archive previous month's consultants",
        value=True,
        help="When the target month differs from the template month, copy the previous AI consultant section to AI - TDS once.",
    )

if template_file is None or attendance_file is None:
    st.info("Upload the payroll template and the monthly attendance summary to begin.")
    st.stop()

# Stable content hashes prevent stale Streamlit state even when filename and size stay unchanged.
template_hash = sha256_bytes(template_file)
attendance_hash = sha256_bytes(attendance_file)
punch_hash = sha256_bytes(punch_file) if punch_file is not None else ""
reset_for_inputs(template_hash, attendance_hash, punch_hash)

# Parse attendance.
try:
    month_from_attendance, attendance_records = parse_attendance_input(
        attendance_file,
        filename=attendance_file.name,
        punch_source=punch_file,
        punch_filename=getattr(punch_file, "name", None) if punch_file is not None else None,
    )
except Exception as exc:
    st.error(f"Could not read attendance workbook: {exc}")
    st.stop()

issues = validate_attendance_records(attendance_records)
if issues:
    st.error("Attendance validation failed:")
    for issue in issues[:20]:
        st.write(f"• {issue}")
    st.stop()

st.session_state.attendance_records = attendance_records

# Inspect template and determine active month/sheets.
try:
    template_wb = load_workbook(template_file)
    template_month = detect_template_month_safe(template_wb)
    active_sheets = detect_active_sheet_names(template_wb, template_month)
    template_data = {
        sheet_name: analyze_template_sheet(template_wb[sheet_name])
        for sheet_name in active_sheets
    }
    identity_register = load_identity_register(template_wb)
    template_wb.close()
except Exception as exc:
    st.error(f"Could not inspect payroll template: {exc}")
    st.stop()

st.session_state.template_month = template_month
st.session_state.active_sheets = active_sheets
st.session_state.template_data = template_data

# Default target month: attendance inferred month, otherwise template month.
default_month = (
    st.session_state.target_month
    or month_from_attendance
    or template_month
    or "August 2026"
)

st.session_state.target_month = st.text_input(
    "Target Payroll Month & Year",
    value=default_month,
)

total_days = get_days_in_month(st.session_state.target_month)

st.markdown("### Run Summary")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Attendance Records", len(attendance_records))
c2.metric("Active Payroll Sheets", len(active_sheets))
c3.metric("Payroll Rows", sum(len(s.get("rows", [])) for sections in template_data.values() for s in sections))
c4.metric("Days in Month", total_days)
punch_corrections = sum(1 for r in attendance_records if r.get("punch_rule_applied"))
punch_discrepancies = sum(1 for r in attendance_records if r.get("punch_discrepancy"))
c5.metric("Punch Corrections", punch_corrections)

if template_month:
    st.caption(f"Template month detected: {template_month}. Active sheets: {', '.join(active_sheets)}")

if punch_file is None:
    st.warning(
        "No raw punch log was uploaded. The app can use the attendance summary, but it cannot detect a missed Punch In / Punch Out from an aggregate summary alone. Upload the raw punch log each month to apply the missing-punch rule."
    )
elif punch_corrections:
    st.success(
        f"Applied missing-punch rule to {punch_corrections} employee record(s): any day with either Punch In or Punch Out is counted as present."
    )

if punch_discrepancies:
    st.warning(
        f"{punch_discrepancies} employee record(s) have fewer unique punch dates than the attendance summary. The summary value was retained for those records to avoid accidental underpayment. Check the raw punch export for completeness."
    )

# Initialize mappings only when the input set changes or mapping is empty.
if not st.session_state.mapping_state:
    st.session_state.mapping_state = {
        m["mapping_key"]: m
        for m in resolve_template_rows(
            template_data,
            attendance_records,
            identity_register,
        )
    }

mappings = st.session_state.mapping_state
attendance_by_id, _ = build_attendance_indexes(attendance_records)
attendance_options = [normalize_employee_id(r.get("employee_id")) for r in attendance_records if normalize_employee_id(r.get("employee_id"))]
attendance_labels = {
    normalize_employee_id(r.get("employee_id")): f"{normalize_employee_id(r.get('employee_id'))} | {r.get('name')} | {format_days(r.get('days_present'))} days"
    for r in attendance_records
    if normalize_employee_id(r.get("employee_id"))
}

conflicts = validate_mapping_conflicts(list(mappings.values()))
review_count = sum(
    1 for m in mappings.values()
    if m.get("needs_confirmation") and not m.get("confirmed")
)
id_match_count = sum(
    1 for m in mappings.values()
    if m.get("confirmed") and m.get("essl_id") and m.get("status") in {STATUS_REGISTERED, STATUS_REGISTERED_NAME_VARIANT, "USER CONFIRMED", "EXACT NAME MATCH"}
)
no_attendance_count = sum(
    1 for m in mappings.values()
    if m.get("status") in {STATUS_NOT_PRESENT, STATUS_MISSING_ATTENDANCE}
)

mc1, mc2, mc3, mc4 = st.columns(4)
mc1.metric("ID / Exact Matched", id_match_count)
mc2.metric("Needs Review", review_count)
mc3.metric("No Attendance", no_attendance_count)
mc4.metric("ID Conflicts", len(conflicts))

if conflicts:
    st.error("The same ESSL Employee ID is assigned to more than one payroll row. Resolve this before generating.")


tab1, tab2, tab3 = st.tabs(["Identity & Attendance", "Payroll Preview", "Generate"])

with tab1:
    st.subheader("Identity Mapping")
    if identity_register:
        st.success(f"Persistent identity register found: {len(identity_register)} mapping(s).")
    else:
        st.info("First run: exact normalized names are auto matched. Name variants require confirmation and will be stored as persistent ESSL IDs in the generated workbook.")

    for sheet_name, sections in template_data.items():
        st.markdown(f"#### {sheet_name}")
        for section in sections:
            st.markdown(f"**{section['type'].capitalize()}**")
            for emp in section.get("rows", []):
                key = payroll_mapping_key(sheet_name, section["type"], int(emp["row_num"]))
                mapping = mappings[key]
                status = mapping.get("status")
                current_id = normalize_employee_id(mapping.get("essl_id"))

                cols = st.columns([2.3, 2.3, 1.0, 1.5])
                cols[0].markdown(
                    f"**{mapping.get('payroll_name')}**  \n{sheet_name} · row {emp['row_num']}"
                )
                cols[1].markdown(
                    f"Payroll code: `{mapping.get('payroll_code') or 'N/A'}`  \nESSL ID: `{current_id or 'None'}`"
                )
                cols[2].markdown(f"**{format_days(mapping.get('days_present'))}** days")

                if mapping.get("needs_confirmation"):
                    cols[3].warning(status or "REVIEW")
                    values = ["__NONE__"] + attendance_options
                    labels = {"__NONE__": "Select attendance identity"}
                    labels.update(attendance_labels)
                    selected = st.selectbox(
                        "Attendance identity",
                        values,
                        index=(values.index(current_id) if current_id in values else 0),
                        format_func=lambda v: labels.get(v, v),
                        key=f"map_select_{key}",
                    )
                    if selected != "__NONE__":
                        if selected != current_id:
                            rec = attendance_by_id[selected]
                            mapping["essl_id"] = selected
                            mapping["attendance_name"] = rec.get("name", "")
                            mapping["days_present"] = float(rec.get("days_present") or 0)
                            mapping["confirmed"] = False
                            mapping["needs_confirmation"] = True
                            mapping["status"] = STATUS_REVIEW
                            mapping["match_method"] = "user_selected"
                            mapping["note"] = "Attendance identity selected by user. Confirmation required."
                            st.rerun()
                        confirmed = st.checkbox(
                            "Confirm identity",
                            value=bool(mapping.get("confirmed")),
                            key=f"confirm_{key}",
                        )
                        mapping["confirmed"] = confirmed
                        mapping["needs_confirmation"] = not confirmed
                        if confirmed:
                            mapping["status"] = "USER CONFIRMED"
                            mapping["match_method"] = "user_confirmed"
                            mapping["note"] = "Identity explicitly confirmed by user."
                    else:
                        mapping["essl_id"] = ""
                        mapping["attendance_name"] = ""
                        mapping["days_present"] = 0.0
                        mapping["confirmed"] = True
                        mapping["needs_confirmation"] = False
                        mapping["status"] = STATUS_NOT_PRESENT
                        mapping["match_method"] = "no_attendance_match"
                        mapping["note"] = "No attendance identity assigned. Payroll will use 0 days."
                else:
                    if status in {STATUS_REGISTERED, STATUS_REGISTERED_NAME_VARIANT, "EXACT NAME MATCH", "USER CONFIRMED"}:
                        cols[3].success(status)
                    elif status == STATUS_MISSING_ATTENDANCE:
                        cols[3].info("Missing this month")
                    else:
                        cols[3].info(status or "Not present")

                if mapping.get("attendance_name"):
                    st.caption(
                        f"Attendance: {mapping['attendance_name']} | ESSL ID: {mapping.get('essl_id')}"
                    )
                if mapping.get("note"):
                    st.caption(mapping["note"])
                st.markdown("---")

    confirmed_ids = {
        normalize_employee_id(m.get("essl_id"))
        for m in mappings.values()
        if m.get("confirmed") and m.get("essl_id")
    }
    new_hire_ids = {
        normalize_employee_id(h.get("essl_id"))
        for h in st.session_state.new_hires
        if h.get("essl_id")
    }
    used_ids = confirmed_ids | new_hire_ids
    unassigned = [
        r for r in attendance_records
        if normalize_employee_id(r.get("employee_id")) not in used_ids
        and str(r.get("name") or "").strip()
        and str(r.get("name") or "").strip().lower() != str(r.get("employee_id") or "").strip().lower()
        and len(str(r.get("name") or "").strip()) >= 3
    ]

    st.markdown("### Unassigned Attendance")
    if unassigned:
        st.dataframe(
            pd.DataFrame(
                [{
                    "ESSL ID": normalize_employee_id(r.get("employee_id")),
                    "Attendance Name": r.get("name"),
                    "Days Present": r.get("days_present"),
                } for r in unassigned]
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success("All plausible attendance records are assigned or intentionally left outside the payroll template.")

    with st.expander("Add a new hire from attendance"):
        if unassigned:
            ids = [normalize_employee_id(r.get("employee_id")) for r in unassigned]
            labels = {normalize_employee_id(r.get("employee_id")): f"{normalize_employee_id(r.get('employee_id'))} | {r.get('name')} | {format_days(r.get('days_present'))} days" for r in unassigned}
            selected_id = st.selectbox("Attendance identity", ids, format_func=lambda v: labels.get(v, v))
            target_sheet = st.selectbox("Payroll sheet", active_sheets)
            sections_available = [s["type"] for s in template_data[target_sheet]]
            target_type = st.selectbox("Section", sections_available)
            designation = st.text_input("Designation", key="new_designation")
            branch = st.text_input("Branch / Function", key="new_branch")
            gross_salary = st.number_input("Gross Salary", min_value=0.0, step=100.0, key="new_gross")
            if st.button("Add New Hire"):
                record = attendance_by_id[selected_id]
                st.session_state.new_hires.append({
                    "essl_id": selected_id,
                    "attendance_name": record.get("name", ""),
                    "days_present": float(record.get("days_present") or 0),
                    "sheet": target_sheet,
                    "type": target_type,
                    "designation": designation.strip(),
                    "branch": branch.strip(),
                    "gross_salary": float(gross_salary),
                })
                st.success(f"Added {record.get('name')} as a pending new hire.")
                st.rerun()
        else:
            st.info("No unassigned attendance record is available for a new hire.")

    if st.session_state.new_hires:
        st.markdown("### New Hires Pending")
        st.dataframe(
            pd.DataFrame([
                {
                    "ESSL ID": h["essl_id"],
                    "Name": h["attendance_name"],
                    "Sheet": h["sheet"],
                    "Section": h["type"],
                    "Days": h["days_present"],
                    "Gross": h["gross_salary"],
                }
                for h in st.session_state.new_hires
            ]),
            use_container_width=True,
            hide_index=True,
        )

with tab2:
    st.subheader("Payroll Preview")
    preview_rows = []
    for mapping in mappings.values():
        confirmed = bool(mapping.get("confirmed") and mapping.get("essl_id"))
        days = float(mapping.get("days_present") or 0) if confirmed else 0.0
        gross = float(mapping.get("gross_salary") or 0)
        prorated = gross * days / total_days if total_days else 0
        if mapping.get("section_type") == "employee":
            deduction = 200.0
            net = prorated - deduction
            kind = "Employee"
        else:
            deduction = prorated * 0.10
            net = prorated - deduction
            kind = "Freelancer (TDS)"
        preview_rows.append({
            "Sheet": mapping.get("sheet_name"),
            "Type": kind,
            "Name": mapping.get("payroll_name"),
            "ESSL ID": mapping.get("essl_id", ""),
            "Present Days": days,
            "Gross Base": gross,
            "Pro-rated Gross": round(prorated, 2),
            "TDS / Prof Tax": round(deduction, 2),
            "Net Salary": round(net, 2),
        })
    if preview_rows:
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Generate Monthly Payroll")

    unresolved_reviews = [m for m in mappings.values() if m.get("needs_confirmation") and not m.get("confirmed")]
    if unresolved_reviews:
        st.warning(f"{len(unresolved_reviews)} identity review(s) still require confirmation. They will remain unresolved in the workbook.")
    if conflicts:
        st.error("Generation is blocked because of duplicate confirmed ESSL IDs.")

    st.caption(
        "The uploaded template is never overwritten. The generated workbook keeps the payroll structure, updates attendance, repairs formula gaps, preserves historical sheets, writes an Attendance Mapping audit sheet, and stores confirmed ESSL identities in a hidden register."
    )

    can_generate = not conflicts
    if st.button("Generate Payroll Excel", type="primary", disabled=not can_generate):
        with st.spinner("Generating payroll workbook..."):
            try:
                generated, audit_rows, active = generate_payroll_workbook(
                    template_source=template_file,
                    attendance_source=attendance_file,
                    target_month=st.session_state.target_month,
                    mappings_by_key=mappings,
                    new_hires=st.session_state.new_hires,
                    archive_enabled=archive_enabled,
                )
                safe_month = re.sub(r"[^A-Za-z0-9]+", "_", st.session_state.target_month).strip("_")
                st.session_state.generated_file = generated
                st.session_state.generated_filename = f"Payroll_{safe_month}.xlsx"
                st.success(f"Payroll generated successfully for {st.session_state.target_month}.")
                st.info(f"Processed active sheets: {', '.join(active)}")
            except Exception as exc:
                st.error(f"Payroll generation failed: {exc}")

if st.session_state.generated_file:
    st.download_button(
        "Download Generated Payroll Excel",
        data=st.session_state.generated_file,
        file_name=st.session_state.generated_filename or "Payroll_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
