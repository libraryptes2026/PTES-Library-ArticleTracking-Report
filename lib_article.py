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
        
        # Map choice to correct Google Sheet layout name
        if cohort_choice == "Lower Sixth (BE Classes)":
            target_spreadsheet = "Articles_Tracker_DB_BE"
        else:
            target_spreadsheet = "Articles_Tracker_DB_AE"
            
        st.info(f"Connected to live database: **{target_spreadsheet}**")
        
        # 🛰️ FETCH REGISTRY DATA FIRST (To use for both Leaderboard & Form Dropdowns)
        try:
            registry_sheet = client.open(target_spreadsheet).worksheet("Student_Registry")
            registry_data = registry_sheet.get_all_records()
            
            # Build the class registry dict for dropdowns
            class_registry = {}
            for row in registry_data:
                # Clean up column spaces dynamically to capture keys safely
                clean_row = {str(k).strip(): v for k, v in row.items()}
                
                # Use flexible keys to match "Form (" column safely
                form_key = [k for k in clean_row.keys() if "Form" in k]
                form_class = str(clean_row.get(form_key[0], "")).strip() if form_key else ""
                student_name = str(clean_row.get("Student Name", "")).strip()
                
                if form_class and student_name:
                    if form_class not in class_registry:
                        class_registry[form_class] = []
                    class_registry[form_class].append(student_name)
                    
            # --- 🏆 TOP 3 READERS LEADERBOARD ENGINE FROM REGISTRY ---
            with st.expander("🏆 View Current Top 3 Reading Leaderboard", expanded=False):
                st.markdown("### 🥇 Top 3 Readers of this Cohort")
                if registry_data:
                    df_reg = pd.DataFrame(registry_data)
                    # Clean trailing whitespaces from headers
                    df_reg.columns = [str(col).strip() for col in df_reg.columns]
                    
                    # Identify Form Class column flexible header mapping
                    f_col = [c for c in df_reg.columns if "Form" in c][0]
                    
                    if "Student Name" in df_reg.columns and "TOTAL" in df_reg.columns:
                        # Enforce numeric typing on Totals column
                        df_reg["TOTAL"] = pd.to_numeric(df_reg["TOTAL"], errors='coerce').fillna(0)
                        
                        # Sort to find the highest total scores
                        top_readers = df_reg.sort_values(by="TOTAL", ascending=False).head(3)
                        
                        if not top_readers.empty and top_readers["TOTAL"].max() > 0:
                            cols_dash = st.columns(3)
                            medals = ["🥇 1st Place", "🥈 2nd Place", "🥉 3rd Place"]
                            
                            for idx, row in enumerate(top_readers.itertuples(index=False)):
                                if idx < len(cols_dash):
                                    with cols_dash[idx]:
                                        st.metric(
                                            label=f"{medals[idx]} ({getattr(row, f_col)})",
                                            value=f"{getattr(row, 'Student Name')}",
                                            delta=f"{int(getattr(row, 'TOTAL'))} Total Articles"
                                        )
                            
                            st.markdown("#### Detailed Leaderboard View")
                            display_df = top_readers[[f_col, "Student Name", "TOTAL"]].rename(
                                columns={f_col: "Form Class", "TOTAL": "Total Articles Read"}
                            )
                            st.dataframe(display_df, use_container_width=True, hide_index=True)
                        else:
                            st.info("No student reading records have totaled totals above 0 yet.")
                    else:
                        st.warning("⚠️ Column headers 'Student Name' or 'TOTAL' were not found in 'Student_Registry'.")
                else:
                    st.info("Student registry is currently empty.")
                    
        except Exception as e:
            # Safe fallback if connection fails completely
            if "BE" in target_spreadsheet:
                class_registry = {"BE1": ["Ahmad Ali", "Dayang Siti"], "BE2": ["Chong Wei", "Nur Huda"]}
            else:
                class_registry = {"AE1": ["Siti Aminah", "Mohammad Noor"], "AE2": ["Aliah Razak", "Khairul Amin"]}
            st.warning(f"⚠️ Note: Could not load live leaderboard registry from '{target_spreadsheet}'. Showing offline entry forms.")

        st.markdown("---")

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
