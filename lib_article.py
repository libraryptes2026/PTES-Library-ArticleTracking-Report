import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

# --- 1. PAGE CONFIGURATION & STYLING ---
st.set_page_config(
    page_title="PTES Library Reading Articles",
    page_icon="📚",
    layout="wide"
)

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
    st.title("PTES Reading Articles")
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

# --- 4. MAIN INTERFACE LOGIC ---
st.title("📚 PTES Library Reading Articles HUB")
st.markdown("Automating the tracking of student reading habits cleanly and efficiently.")
st.markdown("---")

if client:
    # 🔐 LIBRARIAN VERIFICATION LAYER
    admin_pass = st.text_input("🗝️ Enter Librarian Credentials to Unlock Tracker Form:", type="password")
    
    if admin_pass == st.secrets["admin_password"]:
        st.success("🔓 Librarian Access Granted.")
        st.markdown("---")
        
        # 🗂️ COHORT GATEWAY DATABASE SELECTOR
        st.markdown("### 🗂️ Step 1: Select Cohort Level")
        cohort_choice = st.selectbox(
            "Choose Student Cohort Database to open:",
            ["Lower Sixth (BE Classes)", "Upper Sixth (AE Classes)"]
        )
        
        if cohort_choice == "Lower Sixth (BE Classes)":
            target_spreadsheet = "Articles_Tracker_DB_BE"
        else:
            target_spreadsheet = "Articles_Tracker_DB_AE"
            
        st.info(f"Connected to live database: **{target_spreadsheet}**")
        
        # --- 🏆 TOP 3 READERS LEADERBOARD ENGINE FROM BOTTOM FORMULA RANGE ---
        with st.expander("🏆 View Current Top 3 Reading Leaderboard", expanded=False):
            st.markdown("### 🥇 Top 3 Readers of this Cohort")
            try:
                registry_sheet = client.open(target_spreadsheet).worksheet("Student_Registry")
                top_cells = registry_sheet.get("N225:P227")
                
                has_content = False
                if top_cells:
                    for row in top_cells:
                        if len(row) > 0 and str(row[0]).strip() != "" and str(row[0]).strip() != "0":
                            has_content = True
                
                if has_content:
                    cols_dash = st.columns(len(top_cells))
                    medals = ["🥇 1st Place", "🥈 2nd Place", "🥉 3rd Place"]
                    
                    for idx, data_row in enumerate(top_cells):
                        if len(data_row) >= 3 and str(data_row[0]).strip():
                            s_name = data_row[0]
                            f_class = data_row[1]
                            t_score = data_row[2]
                            
                            with cols_dash[idx]:
                                st.metric(
                                    label=f"{medals[idx]} ({f_class})",
                                    value=f"{s_name}",
                                    delta=f"{t_score} Total Articles"
                                )
                    
                    leaderboard_data = []
                    for data_row in top_cells:
                        if len(data_row) >= 3 and str(data_row[0]).strip():
                            leaderboard_data.append({
                                "Form Class": data_row[1],
                                "Student Name": data_row[0],
                                "Total Articles Read": int(data_row[2]) if str(data_row[2]).isdigit() else data_row[2]
                            })
                    if leaderboard_data:
                        st.markdown("#### Detailed Leaderboard View")
                        st.dataframe(pd.DataFrame(leaderboard_data), use_container_width=True, hide_index=True)
                else:
                    st.info("ℹ️ No leaderboard records yet. This cohort has no recorded reading counts higher than 0.")
            except Exception as e:
                st.error(f"Could not extract leaderboard formula cells: {e}")
                
        st.markdown("---")

        # 🛰️ FETCH REGISTRY DYNAMICALLY (SINGLE CELL REQUEST OUTSMARTS THE 429 QUOTA LIMIT)
        class_registry = {}
        try:
            registry_sheet = client.open(target_spreadsheet).worksheet("Student_Registry")
            
            # Pull rows 2 to 210 for columns A & B all in 1 combined query block
            registry_block = registry_sheet.get("A2:B210")
            
            if registry_block:
                for row in registry_block:
                    if len(row) >= 2:
                        f_class_clean = str(row[0]).strip()
                        s_name_clean = str(row[1]).strip()
                        
                        if f_class_clean and s_name_clean and "Form" not in f_class_clean:
                            if f_class_clean not in class_registry:
                                class_registry[f_class_clean] = []
                            class_registry[f_class_clean].append(s_name_clean)
        except Exception as e:
            st.warning(f"⚠️ API Quota Limit encountered. Retrying connection momentarily... Details: {e}")
            
        # Hardcoded dynamic fallback array layout structures
        if not class_registry:
            if "BE" in target_spreadsheet:
                class_registry = {"BE 1": ["Ahmad Ali"], "BE 2": ["Chong Wei"]}
            else:
                class_registry = {"AE 1": ["AE Class Student Registry Empty"], "AE 2": ["No active students registered yet"]}

        # 📝 ENTRY DETAILS FIELD
        st.markdown("### 📝 Step 2: Enter Student Tally Details")
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
        
        # 🗓️ SMART CALENDAR MONTH ANCHORING LOGIC
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
        
        # OPTIONAL REMARKS INPUT BAR
        student_remarks = st.text_input(
            "📝 6. Additional Remarks / Student Status Notes (Optional):", 
            value="", 
            placeholder="e.g., Transferred to Polytechnic / Special consideration"
        )
        
        st.markdown("---")
        
        # Submit Operation
        if st.button("🔥 Submit Tally Data Logs", use_container_width=True):
            if "Empty" in selected_student or "No active students" in selected_student:
                st.error("Cannot submit data into a placeholder student value.")
            else:
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
                    st.success(f"🎉 Success! Recorded {article_count} articles for {selected_student} ({selected_class}) into **{target_spreadsheet}**.")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"Failed to record data. Please check worksheet connection. Details: {e}")
                
    elif admin_pass != "":
        st.error("❌ Invalid Credentials. Access Denied.")
