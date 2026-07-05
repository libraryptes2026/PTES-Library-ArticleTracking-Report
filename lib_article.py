import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import io
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# --- 1. PAGE CONFIGURATION & STYLING ---
st.set_page_config(
    page_title="PTES Reading Articles Tracker",
    page_icon="📚",
    layout="wide"
)

# --- Helper Function for Word Table Borders ---
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

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
        
        # 🗂️ STEP 1: GLOBAL COHORT SELECTOR (Applies to both entry and reports)
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
            
            # Convert raw values into a clean DataFrame for processing
            headers = all_rows[0]
            data_rows = all_rows[1:]
            df_registry = pd.DataFrame(data_rows, columns=headers)
            
            # Reconstruct class mapping for the logging form dropdowns
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

        # 🧭 CREATING NAVIGATION TABS AT THE TOP
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
                
                # Smart month anchoring logic
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
            st.markdown("Extracts live tracking counts dynamically calculated by your `Student_Registry` Sheet formulas.")
            
            if df_registry.empty:
                st.warning("Waiting for connection to active Google Sheet...")
            else:
                # Class Filtering Dropdown
                available_classes = sorted(list(df_registry["Form Class"].unique()))
                report_class = st.selectbox("Select Target Class to View:", available_classes)
                
                # Identify month lists & determine dynamic cut-off based on current date
                all_months = ["MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUG", "SEPT", "OCT"]
                current_month_name = datetime.now().strftime("%B").upper()
                
                # Fallback to map standard shorthand headers if sheet has short version
                sheet_headers = [h.upper() for h in df_registry.columns]
                
                # Filter rows belonging strictly to selected class
                filtered_df = df_registry[df_registry["Form Class"] == report_class].copy()
                
                # Keep core identity data plus month headers that exist in sheet
                visible_columns = ["Form Class", "Student Name"]
                months_to_include = []
                
                for m in all_months:
                    # Look for exact or partial matches in spreadsheet headers
                    found_col = next((col for col in df_registry.columns if col.upper().startswith(m[:3])), None)
                    if found_col:
                        visible_columns.append(found_col)
                        months_to_include.append(found_col)
                        # Break dynamic display check if we reached current month
                        if m == current_month_name or (m == "AUG" and current_month_name == "AUGUST") or (m == "SEPT" and current_month_name == "SEPTEMBER"):
                            pass # If you want to stop exactly at today's month, uncomment 'break' below
                            # break
                
                display_table = filtered_df[visible_columns].reset_index(drop=True)
                
                # Show elegant styled visual summary grid
                st.dataframe(display_table, use_container_width=True)
                
                # --- WORD FILE GENERATION MOTOR ---
                doc = Document()
                
                # Document Title
                title_p = doc.add_paragraph()
                title_run = title_p.add_run(f"Pusat Tingkatan Enam Sengkurong (PTES)\nReading Articles Tally Summary Report")
                title_run.bold = True
                title_run.font.size = Pt(16)
                title_run.font.name = 'Arial'
                title_p.alignment = 1 # Center
                
                # Metaparameters Subtext
                meta_p = doc.add_paragraph()
                meta_run = meta_p.add_run(f"Cohort: {cohort_choice}  |  Class Target: {report_class}\nGenerated On: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
                meta_run.font.size = Pt(10)
                meta_run.italic = True
                meta_p.alignment = 1
                
                doc.add_paragraph().paragraph_format.space_after = Pt(12)
                
                # Build Word Table structure matching web matrix data
                cols_count = len(visible_columns)
                table = doc.add_table(rows=1, cols=cols_count)
                table.style = 'Light Shading Accent 1'
                
                # Insert Headers with nice sizing
                hdr_cells = table.rows[0].cells
                for idx, col_name in enumerate(visible_columns):
                    hdr_cells[idx].text = str(col_name)
                    hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
                    hdr_cells[idx].paragraphs[0].runs[0].font.size = Pt(10)
                    set_cell_margins(hdr_cells[idx], top=120, bottom=120, left=150, right=150)
                    
                    # Apply a classy navy dark blue fill color to header
                    shading_elm = parse_xml(r'<w:shd {} w:fill="1F497D"/>'.format(nsdecls('w')))
                    hdr_cells[idx]._tc.get_or_add_tcPr().append(shading_elm)
                    hdr_cells[idx].paragraphs[0].runs[0].font.color.rgb = docx.shared.RGBColor(255, 255, 255) if 'docx' in globals() else None
                
                # Populate Data Rows
                for _, row in display_table.iterrows():
                    row_cells = table.add_row().cells
                    for idx, col_name in enumerate(visible_columns):
                        val_str = str(row[col_name])
                        # If count is 0 or empty, display a clean hyphen or 0 cleanly
                        row_cells[idx].text = val_str if val_str.strip() != "" else "0"
                        row_cells[idx].paragraphs[0].runs[0].font.size = Pt(9.5)
                        set_cell_margins(row_cells[idx], top=80, bottom=80, left=150, right=150)
                
                # Save into in-memory stream object for download handling
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                st.markdown("---")
                # 📥 ELEGANT ACTIONABLE DOWNLOAD CONTAINER BUTTON
                st.download_button(
                    label=f"📥 Download {report_class} Reading Summary (.docx)",
                    data=doc_io,
                    file_name=f"PTES_Reading_Report_{report_class}_{datetime.now().strftime('%b%Y')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
    elif admin_pass != "":
        st.error("❌ Invalid Credentials. Access Denied.")
