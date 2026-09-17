
from __future__ import annotations

import difflib
import re
from copy import copy
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

IDENTITY_SHEET_NAME = "__Employee_Identity"

REGISTER_HEADERS = [
    "ESSL Employee ID",
    "Attendance Name",
    "Payroll Sheet",
    "Payroll Section",
    "Payroll Row",
    "Payroll Code",
    "Payroll Name",
    "Match Method",
    "Verified",
    "Active",
    "Last Seen Month",
    "Updated At",
]

STATUS_REGISTERED = "ID MATCH"
STATUS_REGISTERED_NAME_VARIANT = "ID MATCH • NAME VARIANT"
STATUS_EXACT_NAME = "EXACT NAME MATCH"
STATUS_REVIEW = "REVIEW NAME"
STATUS_NOT_PRESENT = "NOT PRESENT"
STATUS_DUPLICATE_NAME = "MULTIPLE ID CANDIDATES"
STATUS_MISSING_ATTENDANCE = "MISSING ATTENDANCE"
STATUS_CONFLICT = "IDENTITY CONFLICT"


def normalize_employee_id(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = text.replace(",", "")
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except (TypeError, ValueError):
        pass
    return text


def normalize_name(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\xa0", " ")
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def name_tokens(value: Any) -> set[str]:
    return set(normalize_name(value).split())


def payroll_mapping_key(
    sheet_name: str,
    section_type: str,
    row_num: int,
) -> str:
    return f"{sheet_name}|{section_type}|{row_num}"


def payroll_name_key(
    sheet_name: str,
    section_type: str,
    payroll_name: str,
) -> str:
    return f"{sheet_name}|{section_type}|{normalize_name(payroll_name)}"


@dataclass
class IdentityCandidate:
    employee_id: str
    attendance_name: str
    days_present: float
    score: float
    reason: str


def _name_similarity(a: str, b: str) -> float:
    na = normalize_name(a)
    nb = normalize_name(b)

    if not na or not nb:
        return 0.0

    if na == nb:
        return 1.0

    ratio = difflib.SequenceMatcher(
        None,
        na,
        nb,
    ).ratio()

    ta = name_tokens(na)
    tb = name_tokens(nb)

    if ta and tb:
        overlap = len(ta & tb) / max(len(ta | tb), 1)
    else:
        overlap = 0.0

    first_a = na.split()[0] if na.split() else ""
    first_b = nb.split()[0] if nb.split() else ""

    first_name_bonus = 0.15 if first_a and first_a == first_b else 0.0

    return min(
        1.0,
        (ratio * 0.55) + (overlap * 0.30) + first_name_bonus,
    )


def ensure_identity_sheet(wb):
    if IDENTITY_SHEET_NAME in wb.sheetnames:
        ws = wb[IDENTITY_SHEET_NAME]
        if ws.max_row == 1 and all(
            ws.cell(1, c).value is None
            for c in range(1, len(REGISTER_HEADERS) + 1)
        ):
            ws.append(REGISTER_HEADERS)
        return ws

    ws = wb.create_sheet(IDENTITY_SHEET_NAME)
    ws.append(REGISTER_HEADERS)

    for cell in ws[1]:
        font = copy(cell.font)
        font.bold = True
        cell.font = font

    ws.sheet_state = "hidden"

    widths = {
        1: 18,
        2: 28,
        3: 24,
        4: 16,
        5: 14,
        6: 16,
        7: 28,
        8: 22,
        9: 12,
        10: 10,
        11: 18,
        12: 22,
    }

    for col_idx, width in widths.items():
        ws.column_dimensions[
            __import__("openpyxl").utils.get_column_letter(col_idx)
        ].width = width

    ws.freeze_panes = "A2"
    return ws


def load_identity_register(wb) -> List[Dict[str, Any]]:
    if IDENTITY_SHEET_NAME not in wb.sheetnames:
        return []

    ws = wb[IDENTITY_SHEET_NAME]

    headers = [
        ws.cell(1, c).value
        for c in range(1, ws.max_column + 1)
    ]

    header_map = {
        str(value).strip(): idx
        for idx, value in enumerate(headers)
        if value is not None
    }

    rows: List[Dict[str, Any]] = []

    for row_idx in range(2, ws.max_row + 1):
        if not any(
            ws.cell(row_idx, c).value is not None
            for c in range(1, ws.max_column + 1)
        ):
            continue

        record: Dict[str, Any] = {}

        for header in REGISTER_HEADERS:
            idx = header_map.get(header)

            if idx is None:
                record[header] = None
            else:
                record[header] = ws.cell(
                    row_idx,
                    idx + 1,
                ).value

        record["_row_num"] = row_idx
        rows.append(record)

    return rows


def _register_indexes(
    register: Iterable[Dict[str, Any]],
) -> Tuple[
    Dict[str, Dict[str, Any]],
    Dict[str, Dict[str, Any]],
    Dict[str, List[Dict[str, Any]]],
]:
    by_mapping_key: Dict[str, Dict[str, Any]] = {}
    by_name_key: Dict[str, Dict[str, Any]] = {}
    by_essl_id: Dict[str, List[Dict[str, Any]]] = defaultdict(list)

    for entry in register:
        if str(entry.get("Active", "Yes")).strip().lower() in {
            "no",
            "false",
            "0",
        }:
            continue

        sheet = str(
            entry.get("Payroll Sheet") or ""
        ).strip()

        section = str(
            entry.get("Payroll Section") or ""
        ).strip()

        row_raw = entry.get("Payroll Row")

        try:
            row_num = int(float(row_raw))
        except (TypeError, ValueError):
            row_num = 0

        payroll_name = str(
            entry.get("Payroll Name") or ""
        ).strip()

        if sheet and section and row_num:
            by_mapping_key[
                payroll_mapping_key(
                    sheet,
                    section,
                    row_num,
                )
            ] = entry

        if sheet and section and payroll_name:
            by_name_key[
                payroll_name_key(
                    sheet,
                    section,
                    payroll_name,
                )
            ] = entry

        essl_id = normalize_employee_id(
            entry.get("ESSL Employee ID")
        )

        if essl_id:
            by_essl_id[essl_id].append(
                entry
            )

    return (
        by_mapping_key,
        by_name_key,
        dict(by_essl_id),
    )


def upsert_identity_register(
    wb,
    mappings: Iterable[Dict[str, Any]],
    payroll_month: Optional[str] = None,
) -> None:
    ws = ensure_identity_sheet(wb)
    register = load_identity_register(wb)

    (
        by_mapping_key,
        by_name_key,
        _,
    ) = _register_indexes(register)

    header_to_col = {
        header: idx + 1
        for idx, header in enumerate(
            REGISTER_HEADERS
        )
    }

    now = datetime.utcnow().replace(
        microsecond=0
    ).isoformat()

    for mapping in mappings:
        essl_id = normalize_employee_id(
            mapping.get("essl_id")
        )

        if not essl_id:
            continue

        sheet = str(
            mapping.get("sheet_name") or ""
        ).strip()

        section = str(
            mapping.get("section_type") or ""
        ).strip()

        row_num = int(
            mapping.get("row_num") or 0
        )

        payroll_code = str(
            mapping.get("payroll_code") or ""
        ).strip()

        payroll_name = str(
            mapping.get("payroll_name") or ""
        ).strip()

        attendance_name = str(
            mapping.get("attendance_name") or ""
        ).strip()

        match_method = str(
            mapping.get("match_method")
            or mapping.get("status")
            or "user_confirmed"
        )

        key = payroll_mapping_key(
            sheet,
            section,
            row_num,
        )

        existing = by_mapping_key.get(
            key
        )

        if existing is None:
            existing = by_name_key.get(
                payroll_name_key(
                    sheet,
                    section,
                    payroll_name,
                )
            )

        if existing is not None:
            target_row = int(
                existing["_row_num"]
            )
        else:
            target_row = ws.max_row + 1

        values = {
            "ESSL Employee ID": essl_id,
            "Attendance Name": attendance_name,
            "Payroll Sheet": sheet,
            "Payroll Section": section,
            "Payroll Row": row_num,
            "Payroll Code": payroll_code,
            "Payroll Name": payroll_name,
            "Match Method": match_method,
            "Verified": "Yes"
                if mapping.get("confirmed", True)
                else "No",
            "Active": "Yes",
            "Last Seen Month": payroll_month or "",
            "Updated At": now,
        }

        for header, value in values.items():
            ws.cell(
                target_row,
                header_to_col[header],
            ).value = value


def build_attendance_indexes(
    attendance_records: Iterable[Dict[str, Any]],
):
    by_id: Dict[str, Dict[str, Any]] = {}
    by_name: Dict[str, List[Dict[str, Any]]] = defaultdict(list)

    for record in attendance_records:
        employee_id = normalize_employee_id(
            record.get("employee_id")
        )

        if not employee_id:
            continue

        enriched = dict(record)
        enriched["employee_id"] = employee_id

        by_id[employee_id] = enriched

        key = normalize_name(
            record.get("name")
        )

        if key:
            by_name[key].append(
                enriched
            )

    return by_id, dict(by_name)


def _candidate_matches(
    payroll_name: str,
    attendance_records: List[Dict[str, Any]],
) -> List[IdentityCandidate]:
    candidates: List[IdentityCandidate] = []

    for record in attendance_records:
        attendance_name = str(
            record.get("name") or ""
        ).strip()

        score = _name_similarity(
            payroll_name,
            attendance_name,
        )

        np = normalize_name(
            payroll_name
        )
        na = normalize_name(
            attendance_name
        )

        reason = "name similarity"

        if np == na:
            reason = "exact normalized name"
        elif (
            np.split()[0:1]
            and np.split()[0:1]
            == na.split()[0:1]
        ):
            reason = "same first name + similarity"

        candidates.append(
            IdentityCandidate(
                employee_id=normalize_employee_id(
                    record.get("employee_id")
                ),
                attendance_name=attendance_name,
                days_present=float(
                    record.get("days_present") or 0
                ),
                score=score,
                reason=reason,
            )
        )

    candidates.sort(
        key=lambda item: item.score,
        reverse=True,
    )

    return candidates


def resolve_template_rows(
    template_data: Dict[str, List[Dict[str, Any]]],
    attendance_records: List[Dict[str, Any]],
    register: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """
    Resolve payroll template rows against ESSL attendance.

    Identity priority:
        1. Existing payroll-row registration
        2. Existing payroll-name registration
        3. Exact unique normalized name
        4. Suggested name similarity requiring confirmation
        5. Not present
    """

    (
        attendance_by_id,
        attendance_by_name,
    ) = build_attendance_indexes(
        attendance_records
    )

    (
        register_by_mapping,
        register_by_name,
        register_by_id,
    ) = _register_indexes(
        register
    )

    all_attendance = list(
        attendance_by_id.values()
    )

    results: List[Dict[str, Any]] = []

    for sheet_name, sections in template_data.items():

        for section in sections:

            section_type = section[
                "type"
            ]

            for row in section.get(
                "rows",
                [],
            ):

                row_num = int(
                    row["row_num"]
                )

                payroll_name = str(
                    row.get("name") or ""
                ).strip()

                key = payroll_mapping_key(
                    sheet_name,
                    section_type,
                    row_num,
                )

                name_key = payroll_name_key(
                    sheet_name,
                    section_type,
                    payroll_name,
                )

                base = {
                    "mapping_key": key,
                    "sheet_name": sheet_name,
                    "section_type": section_type,
                    "row_num": row_num,
                    "payroll_code": row.get(
                        "payroll_code"
                    ),
                    "payroll_name": payroll_name,
                    "gross_salary": row.get(
                        "gross_salary"
                    ),
                    "essl_id": "",
                    "attendance_name": "",
                    "days_present": 0.0,
                    "status": STATUS_NOT_PRESENT,
                    "match_method": "",
                    "confirmed": False,
                    "needs_confirmation": False,
                    "suggested_id": "",
                    "suggested_name": "",
                    "note": "",
                }

                # ----------------------------------------------------------
                # 1. Registered by exact payroll row
                # ----------------------------------------------------------
                registered = register_by_mapping.get(
                    key
                )

                if registered is None:
                    registered = register_by_name.get(
                        name_key
                    )

                if registered is not None:
                    registered_id = normalize_employee_id(
                        registered.get(
                            "ESSL Employee ID"
                        )
                    )

                    attendance = (
                        attendance_by_id.get(
                            registered_id
                        )
                        if registered_id
                        else None
                    )

                    base["essl_id"] = registered_id

                    if attendance:
                        base["attendance_name"] = str(
                            attendance.get("name") or ""
                        )

                        base["days_present"] = float(
                            attendance.get(
                                "days_present",
                                0,
                            )
                            or 0
                        )

                        registered_name = str(
                            registered.get(
                                "Attendance Name"
                            )
                            or ""
                        )

                        current_name = str(
                            attendance.get(
                                "name"
                            )
                            or ""
                        )

                        if (
                            normalize_name(
                                registered_name
                            )
                            == normalize_name(
                                current_name
                            )
                        ):
                            base["status"] = (
                                STATUS_REGISTERED
                            )
                            base["note"] = (
                                "Matched using the "
                                "persistent ESSL ID."
                            )
                        else:
                            base["status"] = (
                                STATUS_REGISTERED_NAME_VARIANT
                            )
                            base["note"] = (
                                "ESSL ID matched. "
                                "Attendance name differs "
                                "from the previously registered "
                                "name."
                            )

                        base["match_method"] = (
                            "registered_id"
                        )
                        base["confirmed"] = True
                        base["needs_confirmation"] = False

                    else:
                        base["status"] = (
                            STATUS_MISSING_ATTENDANCE
                        )
                        base["match_method"] = (
                            "registered_id_missing_this_month"
                        )
                        base["confirmed"] = True
                        base["needs_confirmation"] = False
                        base["note"] = (
                            "Employee is registered in the "
                            "payroll identity register but the "
                            "ESSL ID is absent from this month's "
                            "attendance file."
                        )

                    results.append(base)
                    continue

                # ----------------------------------------------------------
                # 2. Exact unique normalized name
                # ----------------------------------------------------------
                exact_candidates = attendance_by_name.get(
                    normalize_name(payroll_name),
                    [],
                )

                if len(exact_candidates) == 1:
                    attendance = exact_candidates[0]

                    base["essl_id"] = normalize_employee_id(
                        attendance.get(
                            "employee_id"
                        )
                    )

                    base["attendance_name"] = str(
                        attendance.get("name") or ""
                    )

                    base["days_present"] = float(
                        attendance.get(
                            "days_present",
                            0,
                        )
                        or 0
                    )

                    base["status"] = (
                        STATUS_EXACT_NAME
                    )

                    base["match_method"] = (
                        "exact_name_first_run"
                    )

                    # Exact unique normalized names are deterministic.
                    base["confirmed"] = True
                    base["needs_confirmation"] = False

                    base["note"] = (
                        "First-run exact normalized "
                        "name match. This ID will be "
                        "persisted after generation."
                    )

                    results.append(base)
                    continue

                if len(exact_candidates) > 1:
                    base["status"] = (
                        STATUS_DUPLICATE_NAME
                    )

                    base["match_method"] = (
                        "duplicate_name"
                    )

                    base["needs_confirmation"] = True

                    base["note"] = (
                        "Multiple ESSL IDs have the "
                        "same normalized name."
                    )

                    results.append(base)
                    continue

                # ----------------------------------------------------------
                # 3. Suggest similarity match
                # ----------------------------------------------------------
                candidates = _candidate_matches(
                    payroll_name,
                    all_attendance,
                )

                best = (
                    candidates[0]
                    if candidates
                    else None
                )

                second = (
                    candidates[1]
                    if len(candidates) > 1
                    else None
                )

                if best and (
                    best.score >= 0.50
                    or (
                        normalize_name(
                            payroll_name
                        ).split()[:1]
                        == normalize_name(
                            best.attendance_name
                        ).split()[:1]
                    )
                ):
                    # Only propose a suggestion.
                    # Never silently treat it as confirmed.
                    base["suggested_id"] = (
                        best.employee_id
                    )

                    base["suggested_name"] = (
                        best.attendance_name
                    )

                    base["essl_id"] = (
                        best.employee_id
                    )

                    base["attendance_name"] = (
                        best.attendance_name
                    )

                    base["days_present"] = (
                        best.days_present
                    )

                    base["status"] = (
                        STATUS_REVIEW
                    )

                    base["match_method"] = (
                        "name_similarity_suggestion"
                    )

                    base["needs_confirmation"] = True
                    base["confirmed"] = False

                    margin = (
                        best.score
                        - second.score
                        if second
                        else best.score
                    )

                    base["note"] = (
                        f"Suggested by name similarity "
                        f"({best.score:.0%}). "
                        f"Confirmation required."
                    )

                    if margin < 0.08 and second:
                        base["note"] += (
                            " Another attendance name is "
                            "also close, so review carefully."
                        )

                    results.append(base)
                    continue

                # ----------------------------------------------------------
                # 4. No attendance record
                # ----------------------------------------------------------
                base["status"] = (
                    STATUS_NOT_PRESENT
                )

                base["match_method"] = (
                    "no_attendance_match"
                )

                base["confirmed"] = True
                base["needs_confirmation"] = False

                base["note"] = (
                    "No reliable ESSL identity match "
                    "was found. Payroll attendance will be 0 "
                    "unless an identity is assigned."
                )

                results.append(base)

    return results


def validate_mapping_conflicts(
    mappings: Iterable[Dict[str, Any]],
) -> Dict[str, List[str]]:
    by_essl: Dict[str, List[str]] = defaultdict(list)

    for mapping in mappings:
        if not mapping.get("essl_id"):
            continue

        if not mapping.get("confirmed"):
            continue

        by_essl[
            normalize_employee_id(
                mapping["essl_id"]
            )
        ].append(
            mapping["mapping_key"]
        )

    duplicates = {
        essl_id: keys
        for essl_id, keys in by_essl.items()
        if len(keys) > 1
    }

    return duplicates
