import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import io
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- 1. PAGE CONFIGURATION & STYLING ---
st.set_page_config(
    page_title="PTES Reading Articles Tracker",
    page_icon="📚",
    layout="wide"
)

# --- Helper Function to Add Page Number Field to Word Header ---
def add_page_number_to_header(header):
    header_p = header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    header_p._p.append(fldChar1)
    header_p._p.append(instrText)
    header_p._p.append(fldChar2)
    header_p._p.append(fldChar3)

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
                current_month_idx = datetime.now().month
                
                visible_columns = ["Form Class", "Student Name"]
                
                for m_name in ordered_months:
                    found_col = next((col for col in df_registry.columns if col.upper().startswith(m_name[:3])), None)
                    if found_col:
                        month_num = month_map.get(m_name.capitalize(), 12)
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
                doc = Document()
                
                # Setup Global Structural Specifications (Letter Size, Landscape, 0.5" Margins)
                for section in doc.sections:
                    section.orientation = WD_ORIENT.LANDSCAPE
                    section.page_width = Inches(11.0)
                    section.page_height = Inches(8.5)
                    
                    # Set 0.5 margins perfectly across all 4 corners
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.5)
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)
                    
                    # Insert dynamic page number field at top center
                    add_page_number_to_header(section.header)
                
                # Document Title Header Block
                title_p = doc.add_paragraph()
                title_p.paragraph_format.space_before = Pt(0)
                title_p.paragraph_format.space_after = Pt(2)
                title_run = title_p.add_run("PUSAT TINGKATAN ENAM SENGKURONG\nSTUDENT READING ARTICLES SUMMARY REPORT")
                title_run.bold = True
                title_run.font.size = Pt(12)
                title_run.font.name = 'Arial'
                title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                meta_p = doc.add_paragraph()
                meta_p.paragraph_format.space_after = Pt(12)
                meta_run = meta_p.add_run(f"CLASS: {report_class}  |  DATABASE COHORT: {cohort_choice.upper()}\nREPORT GENERATED ON: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
                meta_run.font.size = Pt(9.5)
                meta_run.italic = True
                meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                word_cols = list(display_table.columns)
                
                # 'Table Grid' baseline style provides solid clean borders out-of-the-box safely
                table = doc.add_table(rows=1, cols=len(word_cols), style='Table Grid')
                table.autofit = True
                
                # Header formatting: CENTERED, ALL-CAPS, BOLD, BLACK TEXT
                hdr_cells = table.rows[0].cells
                for idx, col_name in enumerate(word_cols):
                    hdr_cells[idx].text = str(col_name).upper()
                    p = hdr_cells[idx].paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Native vertical padding simulation via space configurations
                    p.paragraph_format.space_before = Pt(6)
                    p.paragraph_format.space_after = Pt(6)
                    
                    run = p.runs[0]
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run.font.name = 'Arial'
                    run.font.color.rgb = RGBColor(0, 0, 0)
                
                # Table rows data filling
                for _, row in display_table.iterrows():
                    row_cells = table.add_row().cells
                    for idx, col_name in enumerate(word_cols):
                        val_str = str(row[col_name])
                        row_cells[idx].text = val_str
                        
                        p = row_cells[idx].paragraphs[0]
                        p.paragraph_format.space_before = Pt(4)
                        p.paragraph_format.space_after = Pt(4)
                        
                        # Content Alignment Rules
                        if col_name not in ["Form Class", "Student Name"]:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            
                        run = p.runs[0]
                        run.font.size = Pt(9.5)
                        run.font.name = 'Arial'
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        
                        if col_name == "TOTAL":
                            run.font.bold = True
                
                # Compress into file-stream memory buffer
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                st.markdown("---")
                st.download_button(
                    label=f"📥 Download Official Landscape {report_class} Summary Document (.docx)",
                    data=doc_io,
                    file_name=f"PTES_Reading_Report_{report_class}_{datetime.now().strftime('%b%Y')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
elif admin_pass != "":
    st.error("❌ Invalid Credentials. Access Denied.")
