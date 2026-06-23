import openpyxl
import pandas as pd
import collections
import re
import calendar
from copy import copy

def parse_attendance_input(file_content):
    """
    Parses a simple input Excel file containing:
    - Month/year header (e.g. 'May 2026')
    - A table of names and days present.
    Returns:
    - month_year (str or None)
    - records (list of dicts with 'name' and 'days_present')
    """
    # Read the first sheet
    df = pd.read_excel(file_content, header=None)
    
    # 1. Search for month/year in the sheet
    month_year = None
    month_pattern = re.compile(
        r'\b(january|february|march|april|may|june|july|august|september|october|november|december)\b\s*,?\s*\b(20\d{2})\b',
        re.IGNORECASE
    )
    for r in range(min(15, len(df))):
        for c in range(df.shape[1]):
            val = str(df.iloc[r, c])
            m = month_pattern.search(val)
            if m:
                month_year = f"{m.group(1).capitalize()} {m.group(2)}"
                break
        if month_year:
            break
            
    # 2. Search for headers row
    data_start_row = -1
    name_col_idx = -1
    days_col_idx = -1
    for r in range(len(df)):
        row_vals = [str(x).strip().lower() if pd.notna(x) else "" for x in df.iloc[r].tolist()]
        has_name = any('name' in val or 'employee' in val or 'consultant' in val or 'freelancer' in val for val in row_vals)
        has_days = any('present' in val or 'days' in val or 'attendance' in val or 'working' in val for val in row_vals)
        if has_name and has_days:
            data_start_row = r + 1
            for c, val in enumerate(row_vals):
                if 'name' in val or 'employee' in val or 'consultant' in val or 'freelancer' in val:
                    name_col_idx = c
                elif 'present' in val or 'days' in val or 'attendance' in val or 'working' in val:
                    days_col_idx = c
            break
            
    if name_col_idx == -1 or days_col_idx == -1:
        # Fallback: search for first column with strings, second column with numbers
        name_col_idx = 0
        days_col_idx = 1
        data_start_row = 1
        
    records = []
    for r in range(data_start_row, len(df)):
        name = df.iloc[r, name_col_idx]
        days = df.iloc[r, days_col_idx]
        if pd.notna(name) and str(name).strip():
            name_str = str(name).strip()
            try:
                days_val = float(days)
                if pd.isna(days_val):
                    days_val = 0.0
            except:
                days_val = 0.0
            records.append({"name": name_str, "days_present": days_val})
            
    return month_year, records

def analyze_template_sheet(sheet):
    """
    Scans a template sheet row by row to locate employee and consultant sections.
    Returns a list of sections, each containing row numbers and employee details.
    """
    sections = []
    rows = list(sheet.iter_rows(values_only=True))
    
    current_section = None
    
    for r_idx, row in enumerate(rows):
        row_num = r_idx + 1
        row_str = [str(x).strip().lower() if x is not None else "" for x in row]
        
        is_emp_header = "name of employee" in row_str
        is_cons_header = "name of consultant" in row_str or "name of freelancer" in row_str
        
        if is_emp_header or is_cons_header:
            if current_section:
                current_section["end_row"] = row_num - 2
                sections.append(current_section)
                
            header_col_idx = row_str.index("name of employee") if is_emp_header else row_str.index(
                "name of consultant" if "name of consultant" in row_str else "name of freelancer"
            )
            
            current_section = {
                "type": "employee" if is_emp_header else "consultant",
                "header_row": row_num,
                "name_col": header_col_idx + 1,
                "start_row": row_num + 1,
                "rows": []
            }
            
            # Find data rows downwards until we hit a TOTAL row
            for sub_r in range(row_num + 1, len(rows) + 1):
                sub_row = rows[sub_r - 1]
                sr_no = sub_row[0]
                name = sub_row[current_section["name_col"] - 1]
                
                # Check for TOTAL row
                first_few = [str(x).upper() for x in sub_row[:4] if x is not None]
                is_total = any("TOTAL" in s for s in first_few) or any("TOTAL" in str(x).upper() for x in sub_row if x is not None)
                
                # Check for SUM formula in row if B and A are empty
                is_sum_formula = False
                for cell_val in sub_row:
                    if isinstance(cell_val, str) and cell_val.startswith('='):
                        if 'SUM(' in cell_val.upper():
                            is_sum_formula = True
                            break
                            
                if is_total or (name is None and sr_no is None and is_sum_formula):
                    current_section["end_row"] = sub_r - 1
                    break
                    
                if name is not None and str(name).strip():
                    # Column G (7) is Gross for employee, Column F (6) is Gross for consultant
                    gross_col = 7 if current_section["type"] == "employee" else 6
                    gross_salary = sub_row[gross_col - 1] if len(sub_row) >= gross_col else 0.0
                    
                    current_section["rows"].append({
                        "row_num": sub_r,
                        "name": str(name).strip(),
                        "designation": sub_row[2] if len(sub_row) > 2 else None,
                        "branch": sub_row[3] if len(sub_row) > 3 else None,
                        "gross_salary": gross_salary
                    })
            sections.append(current_section)
            current_section = None
            
    return sections

def get_days_in_month(month_year_str):
    """
    Returns the number of days in the month (e.g. 'May 2026' -> 31)
    """
    try:
        parts = month_year_str.split()
        month_name = parts[0]
        year = int(parts[1])
        month_num = list(calendar.month_name).index(month_name.capitalize())
        return calendar.monthrange(year, month_num)[1]
    except Exception as e:
        print(f"Error calculating days in month: {e}")
        return 30

def update_headers_in_sheet(sheet, old_month_str, new_month_str):
    """
    Replaces occurrences of old month string (e.g. April 2026) with the new month string (e.g. May 2026) in all cells.
    """
    for r in range(1, sheet.max_row + 1):
        for c in range(1, sheet.max_column + 1):
            val = sheet.cell(row=r, column=c).value
            if val and isinstance(val, str):
                if old_month_str.lower() in val.lower():
                    # Case-insensitive replacement
                    pattern = re.compile(re.escape(old_month_str), re.IGNORECASE)
                    new_val = pattern.sub(new_month_str, val)
                    sheet.cell(row=r, column=c).value = new_val
                # Check for other variations like April, 2026
                alt_old = old_month_str.replace(" ", ", ")
                alt_new = new_month_str.replace(" ", ", ")
                if alt_old.lower() in val.lower():
                    pattern = re.compile(re.escape(alt_old), re.IGNORECASE)
                    new_val = pattern.sub(alt_new, val)
                    sheet.cell(row=r, column=c).value = new_val

def archive_previous_month_consultants(wb, old_month_year_str):
    """
    Copies April 2026 consultants from 'AI' sheet to the bottom of the 'AI - TDS' log.
    """
    ai_sheet = wb["AI"]
    tds_sheet = wb["AI - TDS"]
    
    # 1. Locate consultant section in AI
    ai_sections = analyze_template_sheet(ai_sheet)
    consultant_section = None
    for sec in ai_sections:
        if sec["type"] == "consultant":
            consultant_section = sec
            break
            
    if not consultant_section:
        return
        
    # 2. Get last row in AI - TDS
    last_row = tds_sheet.max_row
    while last_row > 1 and not any(tds_sheet.cell(row=last_row, column=c).value for c in range(1, 20)):
        last_row -= 1
        
    start_append_row = last_row + 3
    
    # 3. Title Row
    tds_sheet.cell(row=start_append_row, column=1).value = f"Consultant Fees for the Month of {old_month_year_str}"
    
    # 4. Copy Headers from Top of AI - TDS
    for offset in range(1, 5):
        source_r = 1 + offset # rows 2 to 5
        target_r = start_append_row + offset
        for c in range(1, 26):
            cell = tds_sheet.cell(row=source_r, column=c)
            new_cell = tds_sheet.cell(row=target_r, column=c)
            new_cell.value = cell.value
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format
                
    # 5. Copy Consultants Data and Shift Row References in Formulas
    source_start_row = consultant_section["start_row"]
    source_end_row = consultant_section["end_row"]
    num_consultants = source_end_row - source_start_row + 1
    target_data_start = start_append_row + 5
    
    for idx in range(num_consultants):
        src_r = source_start_row + idx
        tgt_r = target_data_start + idx
        
        for c in range(1, 26):
            src_cell = ai_sheet.cell(row=src_r, column=c)
            tgt_cell = tds_sheet.cell(row=tgt_r, column=c)
            
            val = src_cell.value
            if isinstance(val, str) and val.startswith('='):
                new_val = re.sub(rf'\b({src_r})\b', str(tgt_r), val)
                tgt_cell.value = new_val
            else:
                tgt_cell.value = val
                
            if src_cell.has_style:
                tgt_cell.font = copy(src_cell.font)
                tgt_cell.fill = copy(src_cell.fill)
                tgt_cell.border = copy(src_cell.border)
                tgt_cell.alignment = copy(src_cell.alignment)
                tgt_cell.number_format = src_cell.number_format
                
    # 6. Copy Total Row and Shift SUM Formulas
    tgt_total_row = target_data_start + num_consultants
    src_total_row = consultant_section["end_row"] + 1
    for c in range(1, 26):
        src_cell = ai_sheet.cell(row=src_total_row, column=c)
        tgt_cell = tds_sheet.cell(row=tgt_total_row, column=c)
        
        val = src_cell.value
        if isinstance(val, str) and val.startswith('='):
            new_val = val
            new_val = re.sub(rf'\b({source_start_row})\b', str(target_data_start), new_val)
            new_val = re.sub(rf'\b({source_end_row})\b', str(tgt_total_row - 1), new_val)
            tgt_cell.value = new_val
        else:
            tgt_cell.value = val
            
        if src_cell.has_style:
            tgt_cell.font = copy(src_cell.font)
            tgt_cell.fill = copy(src_cell.fill)
            tgt_cell.border = copy(src_cell.border)
            tgt_cell.alignment = copy(src_cell.alignment)
            tgt_cell.number_format = src_cell.number_format

