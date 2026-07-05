import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import io
import docx
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# --- 1. PAGE CONFIGURATION & STYLING ---
st.set_page_config(
    page_title="PTES Reading Articles Tracker",
    page_icon="📚",
    layout="wide"
)

# --- Helper Function for Pristine Administrative Formatting ---
def format_cell_borders_and_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    
    # 1. Clear any native theme shading styles entirely (Enforce pure white background)
    shading_elm = parse_xml(r'<w:shd {} w:fill="FFFFFF"/>'.format(nsdecls('w')))
    tcPr.append(shading_elm)
    
    # 2. Custom Cell Padding
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)
    
    # 3. Crisp, Bold Solid Dark Borders around every cell
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '6')          # Clean, well-defined border thickness
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')  # Solid sharp black line
        tcBorders.append(border)
    tcPr.append(tcBorders)

# --- 2. SECURE DATABASE CONNECTION (GOOGLE SHEETS) ---
def connect_to_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_info = st.secrets["gspread_creds"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Database Connection Error: {e}")
        return None

client = connect_to_sheets()

# --- 3. SIDEBAR BRANDING & DIGITAL CITIZENSHIP ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c3/Flag_of_Brunei.svg/180px-Flag_of_Brunei.svg.png", width=100)
    st.title("PTES Library Services")
    st.markdown("### 📋 System Guidelines")
    
    with st.expander("📖 Reading Articles Rules"):
        st.write("""
        1. **Article Schedule:** Chief Librarian releases 24-25 articles monthly (Monday–Thursday & Saturday).
        2. **Honesty & Integrity:** Tallies must correspond strictly to verified student readings.
        3. **Librarian Control:** Only authorized library staff may unlock and submit records.
        """)
    
    st.divider()
    st.markdown("""
        <style>
        .footer-line {
            font-size: 11px;
            font-weight: bold;
            text-align: center;
            margin-bottom: 8px;
        }
        .dev-line {
            font-size: 11px;
            text-align: center;
            color: #888;
        }
        </style>
        
        <div class="footer-line">
            ⬜ Perseverance &nbsp; 🟥 Trustworthiness &nbsp; 🔵 Exemplary &nbsp; 🟡 Self-reliance &nbsp; 🟢
        </div>
        <div class="dev-line">
            Web Developer: Miss Hjh Nurul Haziqah HN (PTES Computer Science)
        </div>
    """, unsafe_allow_html=True)

# --- 4. MAIN INTERFACE ---
st.title("📚 PTES Reading Articles Tracker")
st.markdown("Automating the tracking of student reading habits cleanly and efficiently.")
st.markdown("---")

if client:
    # 🔐 LIBRARIAN VERIFICATION LAYER
    admin_pass = st.text_input("🗝️ Enter Librarian Credentials to Unlock Portal:", type="password")
    
    if admin_pass == st.secrets["admin_password"]:
        st.success("🔓 Librarian Access Granted.")
        st.markdown("---")
        
        # 🗂️ STEP 1: GLOBAL COHORT SELECTOR
        st.markdown("### 🗂️ Step 1: Select Cohort Level")
        cohort_choice = st.selectbox(
            "Choose Student Cohort Database to access:",
            ["Lower Sixth (BE Classes)", "Upper Sixth (AE Classes)"]
        )
        
        if cohort_choice == "Lower Sixth (BE Classes)":
            target_spreadsheet = "Articles_Tracker_DB_BE"
        else:
            target_spreadsheet = "Articles_Tracker_DB_AE"
            
        st.info(f"Connected to live database: **{target_spreadsheet}**")
        st.markdown("---")

        # 🛰️ FETCH FULL REGISTRY DATA
        try:
            registry_sheet = client.open(target_spreadsheet).worksheet("Student_Registry")
            all_rows = registry_sheet.get_all_values()
            
            headers = all_rows[0]
            data_rows = all_rows[1:]
            df_registry = pd.DataFrame(data_rows, columns=headers)
            
            class_registry = {}
            for _, row in df_registry.iterrows():
                form_class = str(row.get("Form Class", "")).strip()
                student_name = str(row.get("Student Name", "")).strip()
                if form_class and student_name:
                    if form_class not in class_registry:
                        class_registry[form_class] = []
                    class_registry[form_class].append(student_name)
                    
        except Exception as e:
            st.error(f"Error accessing Google Sheet: {e}")
            df_registry = pd.DataFrame()
            class_registry = {}

        # 🧭 NAVIGATION TABS
        portal_tab1, portal_tab2 = st.tabs(["📝 Submit Tally Logs", "📊 Class Statistics Report"])

        # ==================== TAB 1: LOG INSERTER ====================
        with portal_tab1:
            if not class_registry:
                st.warning("No registry data found. Please check your Google Sheet tabs.")
            else:
                st.markdown("### 📝 Enter Student Tally Details")
                col1, col2 = st.columns(2)
                
                with col1:
                    sorted_classes = sorted(list(class_registry.keys()))
                    selected_class = st.selectbox("1. Select Form Class:", sorted_classes, key="challenge_class")
                    
                    student_options = sorted(class_registry.get(selected_class, []))
                    selected_student = st.selectbox("2. Select Student Name:", student_options, key="challenge_student")
                
                with col2:
                    months_list = ["March", "April", "May", "June", "July", "August", "September", "October"]
                    selected_month = st.selectbox("3. Select Challenge Month:", months_list)
                    article_count = st.number_input("4. Amount of Reading Articles Done:", min_value=1, max_value=31, step=1)
                
                st.markdown("---")
                st.write(f"📂 **5. Enter the specific calendar dates for the {int(article_count)} materials read:**")
                
                month_map = {
                    "March": 3, "April": 4, "May": 5, "June": 6, 
                    "July": 7, "August": 8, "September": 9, "October": 10
                }
                target_month_num = month_map.get(selected_month, datetime.now().month)
                current_year = datetime.now().year
                default_calendar_date = datetime(current_year, target_month_num, 1)
                
                reading_dates = []
                cols = st.columns(4)
                
                for i in range(int(article_count)):
                    col_index = i % 4
                    with cols[col_index]:
                        date_val = st.date_input(
                            f"Article #{i+1} Date", 
                            value=default_calendar_date,
                            key=f"{selected_month}_challenge_date_{i}"
                        )
                        reading_dates.append(date_val.strftime("%d/%m/%Y"))
                
                st.markdown("---")
                student_remarks = st.text_input(
                    "📝 6. Additional Remarks / Student Status Notes (Optional):", 
                    value="", placeholder="e.g., Special consideration"
                )
                
                st.markdown("---")
                
                if st.button("🔥 Submit Tally Data Logs", use_container_width=True):
                    try:
                        challenge_db = client.open(target_spreadsheet).worksheet("Reading_Article_DB")
                        dates_string = ", ".join(reading_dates)
                        
                        new_tally_row = [
                            str(datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
                            str(selected_class),
                            str(selected_student),
                            str(selected_month),
                            str(int(article_count)),
                            str(dates_string),
                            str(student_remarks).strip()
                        ]
                        
                        challenge_db.append_row(new_tally_row, value_input_option="USER_ENTERED")
                        st.success(f"🎉 Success! Recorded {article_count} articles into **{target_spreadsheet}**.")
                        st.balloons()
                    except Exception as e:
                        st.error(f"Failed to record data: {e}")

        # ==================== TAB 2: LIVE REPORT GENERATOR ====================
        with portal_tab2:
            st.markdown("### 📊 Class Reading Overview")
            
            if df_registry.empty:
                st.warning("Waiting for connection to active Google Sheet...")
            else:
                available_classes = sorted(list(df_registry["Form Class"].unique()))
                report_class = st.selectbox("Select Target Class to View:", available_classes)
                
                # Chronological calendar limit check
                ordered_months = ["MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER"]
                current_month_idx = datetime.now().month  # Dynamic evaluation (7 for July)
                
                visible_columns = ["Form Class", "Student Name"]
                
                for m_name in ordered_months:
                    found_col = next((col for col in df_registry.columns if col.upper().startswith(m_name[:3])), None)
                    if found_col:
                        month_num = month_map.get(m_name.capitalize(), 12)
                        # Only include columns up to today's current calendar month (July)
                        if month_num <= current_month_idx:
                            visible_columns.append(found_col)
                
                filtered_df = df_registry[df_registry["Form Class"] == report_class].copy()
                
                display_table = filtered_df[visible_columns].reset_index(drop=True)
                display_table = display_table.replace(["#N/A", "nan", ""], "0")
                
                # Math Processing Engine for Cumulative Sums
                month_cols_only = [c for c in visible_columns if c not in ["Form Class", "Student Name"]]
                for col in month_cols_only:
                    display_table[col] = pd.to_numeric(display_table[col], errors='coerce').fillna(0).astype(int)
                
                # Append TOTAL Summary column on the far right edge
                display_table["TOTAL"] = display_table[month_cols_only].sum(axis=1)
                
                st.dataframe(display_table, use_container_width=True)
                
                # --- DOCUMENT GENERATION ENGINE ---
                #################################
                doc = Document()

                # Loop through sections (or target a specific section like doc.sections[0])
                for section in doc.sections:
                # 1. Toggle the orientation property to landscape
                section.orientation = WD_ORIENT.LANDSCAPE
    
                # 2. Physically swap the dimensions (Standard Letter size example)
                section.page_width = Inches(11.0)
                section.page_height = Inches(8.5)
    
                #################################
                # Set Standard Margins (0.5 inch everywhere)
                for section in doc.sections:
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.5)
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)
                
                # Document Title Header Block
                title_p = doc.add_paragraph()
                title_run = title_p.add_run("PUSAT TINGKATAN ENAM SENGKURONG\nSTUDENT READING ARTICLES SUMMARY REPORT")
                title_run.bold = True
                title_run.font.size = Pt(12)
                title_run.font.name = 'Arial'
                title_run.font.color.rgb = docx.shared.RGBColor(0, 0, 0)
                title_p.alignment = 1
                
                meta_p = doc.add_paragraph()
                meta_run = meta_p.add_run(f"CLASS: {report_class}  |  DATABASE COHORT: {cohort_choice.upper()}\nREPORT GENERATED ON: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
                meta_run.font.size = Pt(9.5)
                meta_run.italic = True
                meta_run.font.color.rgb = docx.shared.RGBColor(0, 0, 0)
                meta_p.alignment = 1
                
                doc.add_paragraph().paragraph_format.space_after = Pt(12)
                
                word_cols = list(display_table.columns)
                table = doc.add_table(rows=1, cols=len(word_cols))
                table.autofit = True
                
                # Header formatting: ALL-CAPS, BOLD, BLACK TEXT, WHITE BACKGROUND
                hdr_cells = table.rows[0].cells
                for idx, col_name in enumerate(word_cols):
                    hdr_cells[idx].text = str(col_name).upper()
                    run = hdr_cells[idx].paragraphs[0].runs[0]
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run.font.name = 'Arial'
                    run.font.color.rgb = docx.shared.RGBColor(0, 0, 0) # Solid Black Text
                    format_cell_borders_and_margins(hdr_cells[idx], top=120, bottom=120, left=150, right=150)
                
                # Table rows data filling
                for _, row in display_table.iterrows():
                    row_cells = table.add_row().cells
                    for idx, col_name in enumerate(word_cols):
                        val_str = str(row[col_name])
                        row_cells[idx].text = val_str
                        
                        run = row_cells[idx].paragraphs[0].runs[0]
                        run.font.size = Pt(9.5)
                        run.font.name = 'Arial'
                        run.font.color.rgb = docx.shared.RGBColor(0, 0, 0)
                        
                        # Emphasize values under the total column
                        if col_name == "TOTAL":
                            run.font.bold = True
                            
                        format_cell_borders_and_margins(row_cells[idx], top=100, bottom=100, left=150, right=150)
                
                # Compress into file-stream memory buffer
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                st.markdown("---")
                st.download_button(
                    label=f"📥 Download Official {report_class} Reading Summary (.docx)",
                    data=doc_io,
                    file_name=f"PTES_Reading_Report_{report_class}_{datetime.now().strftime('%b%Y')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
    elif admin_pass != "":
        st.error("❌ Invalid Credentials. Access Denied.")
