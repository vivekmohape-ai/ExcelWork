import streamlit as st
import pandas as pd
import openpyxl
import os
import io
import difflib
import re
from excel_utils import (
    parse_attendance_input,
    analyze_template_sheet,
    get_days_in_month,
    update_headers_in_sheet,
    archive_previous_month_consultants
)

# Set page configuration with a premium icon and layout
st.set_page_config(
    page_title=" Payroll Automator",
    page_icon="💸💸💸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium CSS for HSL colors, Google Fonts, hover effects, and cards
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global Styles */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* Main Gradient Header */
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
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    
    /* Styled Card Container */
    .premium-card {
        background: rgba(255, 255, 255, 0.7);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(229, 231, 235, 0.5);
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        margin-bottom: 1rem;
    }
    
    /* Custom buttons */
    div.stButton > button {
        background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%);
        color: white;
        border: none;
        padding: 0.6rem 1.8rem;
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 10px rgba(99, 102, 241, 0.3);
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%);
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(99, 102, 241, 0.4);
    }
    
    /* Metric Card Styling */
    .metric-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1e1b4b;
        margin-bottom: 0.2rem;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #4f46e5;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
</style>
""", unsafe_allow_html=True)

# Main Application Layout
st.markdown("<h1 class='header-gradient'>Payroll Automator</h1>", unsafe_allow_html=True)
st.markdown("<p class='subheader-text'>Automated monthly payroll matching, calculations, and styled Excel generation.</p>", unsafe_allow_html=True)

# Define path to default template file
DEFAULT_TEMPLATE_PATH = r"C:\Users\rohit\OneDrive\Documents\Another idea\hr Attendance\files\MERGE FILE - PRINCE - Till APRIL 2026.xlsx"

# Initialize Session State Variables
if "template_data" not in st.session_state:
    st.session_state.template_data = None
if "input_data" not in st.session_state:
    st.session_state.input_data = None
if "parsed_month" not in st.session_state:
    st.session_state.parsed_month = None
if "input_records" not in st.session_state:
    st.session_state.input_records = []
if "name_mappings" not in st.session_state:
    st.session_state.name_mappings = {}
if "new_employees" not in st.session_state:
    st.session_state.new_employees = [] # List of dicts representing manually added employees
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
if "download_filename" not in st.session_state:
    st.session_state.download_filename = None
if "last_attendance_key" not in st.session_state:
    st.session_state.last_attendance_key = None
if "last_template_key" not in st.session_state:
    st.session_state.last_template_key = None

# ----------------- SIDEBAR (CONFIGURATION) -----------------
with st.sidebar:
    st.markdown("### ⚙️ Payroll Configuration")
    
    # 1. Template File
    st.markdown("**1. Select Template Excel File**")
    template_option = st.radio(
        "Template Source",
        ["Use Default Template", "Upload Custom Template"],
        label_visibility="collapsed"
    )
    
    template_file_content = None
    if template_option == "Use Default Template":
        if os.path.exists(DEFAULT_TEMPLATE_PATH):
            template_file_content = DEFAULT_TEMPLATE_PATH
            st.success("Loaded default template.")
        else:
            st.error("Default template not found. Please upload manually.")
    else:
        uploaded_template = st.file_uploader("Upload Merge Template File", type=["xlsx"])
        if uploaded_template:
            template_file_content = uploaded_template
            st.success("Uploaded custom template.")
            
    st.markdown("---")
    
    # 2. Input Attendance File
    st.markdown("**2. Upload Monthly Attendance**")
    uploaded_attendance = st.file_uploader(
        "Upload input Excel with names and days present",
        type=["xlsx", "xls"]
    )
    
    st.markdown("---")
    
    # 3. Archiving Settings
    st.markdown("**3. Historical Log Options**")
    archive_enabled = st.checkbox("Archive previous month's consultants to AI - TDS", value=True)
    
    # 4. Download Center
    if st.session_state.generated_file is not None:
        st.markdown("---")
        st.markdown("### 📥 Download Center")
        st.download_button(
            label="📥 Download Payroll Excel",
            data=st.session_state.generated_file,
            file_name=st.session_state.download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="sidebar_download_button"
        )

# ----------------- DATA LOADING AND PARSING -----------------
if template_file_content and uploaded_attendance:
    # Check for changes in uploaded files to reset states dynamically
    attendance_key = (uploaded_attendance.name, uploaded_attendance.size)
    if attendance_key != st.session_state.last_attendance_key:
        print(f"DEBUG: Resetting state because attendance file changed from {st.session_state.last_attendance_key} to {attendance_key}")
        st.session_state.last_attendance_key = attendance_key
        st.session_state.input_data = None
        st.session_state.parsed_month = None
        st.session_state.input_records = []
        st.session_state.generated_file = None
        st.session_state.name_mappings = {}
        st.session_state.new_employees = []

    if isinstance(template_file_content, str):
        template_key = template_file_content
    else:
        template_key = (template_file_content.name, template_file_content.size)
        
    if template_key != st.session_state.last_template_key:
        print(f"DEBUG: Resetting state because template file changed from {st.session_state.last_template_key} to {template_key}")
        st.session_state.last_template_key = template_key
        st.session_state.template_data = None
        st.session_state.generated_file = None
        st.session_state.name_mappings = {}
        st.session_state.new_employees = []

    # 1. Load the input attendance file and parse names/month
    try:
        if st.session_state.input_data is None or st.session_state.parsed_month is None:
            # Parse only once or when inputs change
            month, records = parse_attendance_input(uploaded_attendance)
            st.session_state.parsed_month = month if month else "May 2026"
            st.session_state.input_records = records
            st.session_state.input_data = True
    except Exception as e:
        st.error(f"Error reading attendance file: {e}")
        st.stop()
        
    # 2. Load template structure
    try:
        if st.session_state.template_data is None:
            wb_temp = openpyxl.load_workbook(template_file_content, read_only=True)
            sheet_sections = {}
            for name in wb_temp.sheetnames:
                if name == "AI - TDS":
                    continue # Skip logs sheet
                sheet = wb_temp[name]
                sheet_sections[name] = analyze_template_sheet(sheet)
            st.session_state.template_data = sheet_sections
            wb_temp.close()
    except Exception as e:
        st.error(f"Error loading template details: {e}")
        st.stop()

    # 3. Global parameters & Month config in the UI
    st.markdown("<div class='premium-card'>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        # Allow user to change parsed month/year
        month_input = st.text_input("Target Payroll Month & Year", value=st.session_state.parsed_month)
        if month_input != st.session_state.parsed_month:
            print(f"DEBUG: Resetting generated file because month input changed from {st.session_state.parsed_month} to {month_input}")
            st.session_state.parsed_month = month_input
            st.session_state.generated_file = None
    with col2:
        total_days = get_days_in_month(st.session_state.parsed_month)
        st.markdown(f"<div class='metric-label'>Total Days in Month</div><div class='metric-value'>{total_days}</div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='metric-label'>Matched Records</div><div class='metric-value'>{len(st.session_state.input_records)}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Prepare lists for mapping UI
    attendance_names = [r["name"] for r in st.session_state.input_records]
    
    # ----------------- TABS SETUP -----------------
    tab1, tab2, tab3 = st.tabs(["👥 Name Mapping & Joiners", "📊 Live Payroll Preview", "💾 Generate & Download Excel"])
    
    # ----------------- TAB 1: NAME MAPPING & JOINERS -----------------
    with tab1:
        st.markdown("### Map Attendance Sheet Names to Template Employees")
        st.info("The application automatically suggests matching names. You can override these using the dropdown selectors below. Select 'Not Present' to set 0 days, or select 'Add as New' if they are new hires.")
        
        # We will loop sheet by sheet and show the matching interface
        template_sheets = st.session_state.template_data
        
        for sheet_name, sections in template_sheets.items():
            if not sections:
                continue
                
            st.markdown(f"#### 🏢 Sheet: `{sheet_name}`")
            
            # Combine all employees and consultants in this sheet
            for section in sections:
                sec_type = "Employee" if section["type"] == "employee" else "Freelancer (TDS)"
                st.markdown(f"**Section: {sec_type}**")
                
                rows_data = []
                for emp in section["rows"]:
                    emp_name = emp["name"]
                    row_num = emp["row_num"]
                    
                    # Compute unique key for the employee row
                    mapping_key = f"{sheet_name}_{section['type']}_{row_num}_{emp_name}"
                    
                    # Automatic recommendation
                    if mapping_key not in st.session_state.name_mappings:
                        # Try to find best match in attendance names
                        best_matches = difflib.get_close_matches(emp_name, attendance_names, n=1, cutoff=0.7)
                        if best_matches:
                            st.session_state.name_mappings[mapping_key] = best_matches[0]
                        else:
                            # Try case-insensitive exact check
                            lower_names = [n.lower() for n in attendance_names]
                            if emp_name.lower() in lower_names:
                                match_idx = lower_names.index(emp_name.lower())
                                st.session_state.name_mappings[mapping_key] = attendance_names[match_idx]
                            else:
                                st.session_state.name_mappings[mapping_key] = "Not Present (0 Days)"
                                
                    current_mapped = st.session_state.name_mappings[mapping_key]
                    
                    # Display the mapping layout
                    cols = st.columns([3, 1, 4])
                    with cols[0]:
                        st.markdown(f"👤 **{emp_name}** `(Template Row {row_num})`  \n*Designation: {emp['designation']} | Gross: {emp['gross_salary']}*")
                    with cols[1]:
                        st.markdown("➡️ maps to")
                    with cols[2]:
                        options = ["Not Present (0 Days)"] + attendance_names
                        selected_mapped = st.selectbox(
                            f"Mapped name for {emp_name} ({row_num})",
                            options=options,
                            index=options.index(current_mapped) if current_mapped in options else 0,
                            key=f"select_{mapping_key}",
                            label_visibility="collapsed"
                        )
                        if selected_mapped != current_mapped:
                            print(f"DEBUG: Resetting generated file because mapping changed for {mapping_key} from {current_mapped} to {selected_mapped}")
                            st.session_state.name_mappings[mapping_key] = selected_mapped
                            st.session_state.generated_file = None
            st.markdown("---")
            
        # Add New Joiners Section
        st.markdown("### ➕ Add New Employee/Freelancer")
        st.markdown("If you have new hires in your input sheet who are not in the template, you can add them to a sheet below. The app will write them into the blank rows in that sheet and generate the formulas.")
        
        with st.form("new_employee_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                new_name = st.selectbox("Select name from input sheet", ["Select Name"] + sorted(attendance_names))
            with col2:
                target_sheet = st.selectbox("Add to Sheet", list(template_sheets.keys()))
            with col3:
                emp_type_choice = st.selectbox("Type", ["Employee", "Freelancer (TDS)"])
                
            col4, col5, col6 = st.columns(3)
            with col4:
                new_designation = st.text_input("Designation")
            with col5:
                new_branch = st.text_input("Branch / Function")
            with col6:
                new_gross = st.number_input("Gross Salary", min_value=0.0, step=100.0)
                
            submitted = st.form_submit_button("Add New Hire")
            if submitted:
                if new_name == "Select Name":
                    st.error("Please select a name from the input list.")
                elif not new_designation or not new_branch:
                    st.error("Please fill in Designation and Branch/Function.")
                else:
                    st.session_state.new_employees.append({
                        "name": new_name,
                        "sheet": target_sheet,
                        "type": "employee" if emp_type_choice == "Employee" else "consultant",
                        "designation": new_designation,
                        "branch": new_branch,
                        "gross_salary": new_gross
                    })
                    st.session_state.generated_file = None
                    st.success(f"Added {new_name} as new {emp_type_choice} in sheet {target_sheet}!")
                    
        # List manually added new employees
        if st.session_state.new_employees:
            st.markdown("##### New Joiners Added:")
            for idx, ne in enumerate(st.session_state.new_employees):
                st.write(f"• **{ne['name']}** ({ne['type'].capitalize()}) in `{ne['sheet']}` — Designation: {ne['designation']}, Gross: {ne['gross_salary']} "
                         f"[ [Remove](javascript:void(0)) ]")
                # Add a clear button for list
            if st.button("Clear All New Joiners"):
                st.session_state.new_employees = []
                st.session_state.generated_file = None
                st.rerun()

    # ----------------- TAB 2: LIVE PAYROLL PREVIEW -----------------
    with tab2:
        st.markdown("### Payroll Calculations Preview")
        st.info("Here is a live preview of the pro-rated calculations before exporting. Loan entries are ignored.")
        
        # Build mapping of input name to days present
        attendance_days = {r["name"]: r["days_present"] for r in st.session_state.input_records}
        
        for sheet_name, sections in template_sheets.items():
            if not sections:
                continue
                
            st.markdown(f"#### 🏢 Sheet: `{sheet_name}`")
            
            preview_rows = []
            
            # Helper to calculate pro-rated columns
            for section in sections:
                for emp in section["rows"]:
                    emp_name = emp["name"]
                    row_num = emp["row_num"]
                    mapping_key = f"{sheet_name}_{section['type']}_{row_num}_{emp_name}"
                    
                    mapped_name = st.session_state.name_mappings.get(mapping_key, "Not Present (0 Days)")
                    days_present = attendance_days.get(mapped_name, 0.0) if mapped_name != "Not Present (0 Days)" else 0.0
                    
                    gross = float(emp["gross_salary"]) if emp["gross_salary"] else 0.0
                    prorated = gross * (days_present / total_days)
                    
                    # Deductions and net
                    if section["type"] == "employee":
                        # Basic=40%, HRA=10%, Conveyance=15%, Misc=25%, Medical=1250, Other=10% - 1250
                        basic = prorated * 0.40
                        hra = prorated * 0.10
                        conveyance = prorated * 0.15
                        misc = prorated * 0.25
                        medical = 1250.0
                        other = (prorated * 0.10) - 1250.0
                        gross_calc = basic + hra + conveyance + misc + medical + other # equals pro-rated gross
                        deductions = 200.0 # Standard Prof Tax
                        net = gross_calc - deductions
                        record_type = "Employee"
                    else:
                        # Freelancer TDS 10%
                        basic = prorated * 0.40
                        hra = prorated * 0.10
                        conveyance = prorated * 0.15
                        misc = prorated * 0.25
                        other = prorated * 0.10
                        actual_gross = basic + hra + conveyance + misc + other
                        deductions = actual_gross * 0.10 # TDS
                        net = actual_gross - deductions
                        record_type = "Freelancer (TDS)"
                        
                    preview_rows.append({
                        "Type": record_type,
                        "Name": emp_name,
                        "Gross Base": gross,
                        "Present Days": days_present,
                        "Pro-rated Gross": round(prorated, 2),
                        "TDS / Prof Tax": round(deductions, 2),
                        "Net Salary": round(net, 2)
                    })
                    
            # Also append manual new employees for preview
            for ne in st.session_state.new_employees:
                if ne["sheet"] == sheet_name:
                    days_present = attendance_days.get(ne["name"], 0.0)
                    gross = float(ne["gross_salary"])
                    prorated = gross * (days_present / total_days)
                    if ne["type"] == "employee":
                        deductions = 200.0
                        net = prorated - deductions
                    else:
                        deductions = prorated * 0.10
                        net = prorated - deductions
                    preview_rows.append({
                        "Type": f"New {ne['type'].capitalize()}",
                        "Name": ne["name"],
                        "Gross Base": gross,
                        "Present Days": days_present,
                        "Pro-rated Gross": round(prorated, 2),
                        "TDS / Prof Tax": round(deductions, 2),
                        "Net Salary": round(net, 2)
                    })
                    
            if preview_rows:
                df_preview = pd.DataFrame(preview_rows)
                st.dataframe(df_preview, use_container_width=True)
            else:
                st.write("No employees or consultants in this sheet.")

    # ----------------- TAB 3: GENERATE & DOWNLOAD -----------------
    with tab3:
        st.markdown("### Generate Monthly Spreadsheet")
        st.write("Click the button below to process your payroll and download the generated styled Excel file.")
        
        # Load the month for title updating
        old_month = "April 2026"
        new_month = st.session_state.parsed_month
        
        st.warning(f"This will replace all references of **'{old_month}'** in sheet titles with **'{new_month}'** and set total days to **{total_days}**.")
        
        if st.button("Generate Payroll Excel"):
            # Load template using openpyxl (data_only=False so we load the formulas)
            with st.spinner("Generating Excel file..."):
                try:
                    # Determine workbook object depending on custom upload or default path
                    if isinstance(template_file_content, str):
                        wb = openpyxl.load_workbook(template_file_content, data_only=False)
                    else:
                        template_file_content.seek(0)
                        wb = openpyxl.load_workbook(template_file_content, data_only=False)
                        
                    # 1. Update month strings in all sheets
                    for s_name in wb.sheetnames:
                        update_headers_in_sheet(wb[s_name], old_month, new_month)
                        
                    # 2. Archive consultants if enabled
                    if archive_enabled and "AI - TDS" in wb.sheetnames:
                        archive_previous_month_consultants(wb, old_month)
                        
                    # Build attendance lookup
                    attendance_days = {r["name"]: r["days_present"] for r in st.session_state.input_records}
                    
                    # 3. Process each sheet's employee and consultant values
                    for sheet_name, sections in st.session_state.template_data.items():
                        sheet = wb[sheet_name]
                        
                        # Process existing employees
                        for section in sections:
                            for emp in section["rows"]:
                                emp_name = emp["name"]
                                row_num = emp["row_num"]
                                mapping_key = f"{sheet_name}_{section['type']}_{row_num}_{emp_name}"
                                
                                mapped_name = st.session_state.name_mappings.get(mapping_key, "Not Present (0 Days)")
                                days_present = attendance_days.get(mapped_name, 0.0) if mapped_name != "Not Present (0 Days)" else 0.0
                                
                                # Write to sheet
                                if section["type"] == "employee":
                                    # Employee: Col E is Total Days, Col F is Days Present
                                    sheet.cell(row=row_num, column=5).value = total_days
                                    sheet.cell(row=row_num, column=6).value = days_present
                                    # Clear Loan/Advance: Col S (19)
                                    sheet.cell(row=row_num, column=19).value = None
                                else:
                                    # Consultant: Col D is Total Days, Col E is Days Present
                                    sheet.cell(row=row_num, column=4).value = total_days
                                    sheet.cell(row=row_num, column=5).value = days_present
                                    # Clear Loan/Advance: Col O (15)
                                    sheet.cell(row=row_num, column=15).value = None
                                    
                        # Process new joiners added to this sheet
                        sheet_new_hires = [ne for ne in st.session_state.new_employees if ne["sheet"] == sheet_name]
                        if sheet_new_hires:
                            for ne in sheet_new_hires:
                                target_sec = None
                                for sec in sections:
                                    if sec["type"] == ne["type"]:
                                        target_sec = sec
                                        break
                                        
                                if target_sec:
                                    empty_row = None
                                    for r_idx in range(target_sec["start_row"], target_sec["end_row"] + 1):
                                        name_val = sheet.cell(row=r_idx, column=target_sec["name_col"]).value
                                        if name_val is None or str(name_val).strip() == "":
                                            empty_row = r_idx
                                            break
                                            
                                    if empty_row:
                                        prev_sr = sheet.cell(row=empty_row - 1, column=1).value
                                        try:
                                            sheet.cell(row=empty_row, column=1).value = int(prev_sr) + 1
                                        except:
                                            sheet.cell(row=empty_row, column=1).value = ""
                                            
                                        sheet.cell(row=empty_row, column=target_sec["name_col"]).value = ne["name"]
                                        sheet.cell(row=empty_row, column=3).value = ne["designation"]
                                        sheet.cell(row=empty_row, column=4).value = ne["branch"]
                                        
                                        days_present = attendance_days.get(ne["name"], 0.0)
                                        
                                        if ne["type"] == "employee":
                                            sheet.cell(row=empty_row, column=5).value = total_days
                                            sheet.cell(row=empty_row, column=6).value = days_present
                                            sheet.cell(row=empty_row, column=7).value = ne["gross_salary"]
                                            sheet.cell(row=empty_row, column=19).value = None
                                            for c in [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 20, 22]:
                                                cell_above = sheet.cell(row=empty_row - 1, column=c)
                                                if cell_above.value and isinstance(cell_above.value, str) and cell_above.value.startswith('='):
                                                    new_f = re.sub(rf'\b({empty_row - 1})\b', str(empty_row), cell_above.value)
                                                    sheet.cell(row=empty_row, column=c).value = new_f
                                                elif c == 12:
                                                    sheet.cell(row=empty_row, column=12).value = 1250
                                                elif c == 15:
                                                    sheet.cell(row=empty_row, column=15).value = 1
                                                elif c == 17:
                                                    sheet.cell(row=empty_row, column=17).value = 200
                                        else:
                                            sheet.cell(row=empty_row, column=4).value = total_days
                                            sheet.cell(row=empty_row, column=5).value = days_present
                                            sheet.cell(row=empty_row, column=6).value = ne["gross_salary"]
                                            sheet.cell(row=empty_row, column=15).value = None
                                            for c in [7, 8, 9, 10, 11, 12, 13, 14, 16]:
                                                cell_above = sheet.cell(row=empty_row - 1, column=c)
                                                if cell_above.value and isinstance(cell_above.value, str) and cell_above.value.startswith('='):
                                                    new_f = re.sub(rf'\b({empty_row - 1})\b', str(empty_row), cell_above.value)
                                                    sheet.cell(row=empty_row, column=c).value = new_f
                                                elif c == 13:
                                                    sheet.cell(row=empty_row, column=13).value = 1
                                    else:
                                        st.warning(f"No blank rows available in `{sheet_name}` for new hire {ne['name']}.")
                                        
                    # Save output workbook to session state bytes
                    buffer = io.BytesIO()
                    wb.save(buffer)
                    wb.close()
                    
                    st.session_state.generated_file = buffer.getvalue()
                    st.session_state.download_filename = f"MERGE FILE - PRINCE - {new_month}.xlsx"
                    st.success("Successfully generated payroll sheet! Click 'Download Generated Excel File' below to save it.")
                except Exception as ex:
                    st.error(f"Error generating payroll Excel: {ex}")
                    
        # Render the download button if the file has been generated
        if st.session_state.generated_file is not None:
            st.info("💡 You can download the generated file below, or from the **Download Center** in the sidebar if the tab refreshes.")
            st.download_button(
                label="📥 Download Generated Excel File",
                data=st.session_state.generated_file,
                file_name=st.session_state.download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="main_download_button"
            )
else:
    st.info("👋 Please upload your Monthly Attendance report and select your Template file in the sidebar to get started.")

