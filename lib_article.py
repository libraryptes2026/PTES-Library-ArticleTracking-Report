import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

# --- 1. PAGE CONFIGURATION & STYLING ---
st.set_page_config(
    page_title="PTES Reading Articles Tracker",
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
    st.title("PTES Library Services")
    st.markdown("### 📋 System Guidelines")
    
    with st.expander("📖 Reading Challenge Rules"):
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
            Web Developer: Miss Hjh Nurul Haziqah HN (Computer Science)
        </div>
    """, unsafe_allow_html=True)

# --- 4. MAIN INTERFACE LOGIC ---
st.title("📚 PTES Annual Reading Articles Tracker")
st.markdown("Automating the tracking of student reading habits cleanly and efficiently.")
st.markdown("---")

if client:
    # 🛰️ FETCH REGISTRY FROM GOOGLE SHEETS
    try:
        registry_sheet = client.open("Articles_Tracker_DB").worksheet("Student_Registry")
        registry_data = registry_sheet.get_all_records()
        
        class_registry = {}
        for row in registry_data:
            form_class = str(row.get("Form Class", "")).strip()
            student_name = str(row.get("Student Name", "")).strip()
            if form_class and student_name:
                if form_class not in class_registry:
                    class_registry[form_class] = []
                class_registry[form_class].append(student_name)
    except Exception:
        class_registry = {
            "BE1": ["Ahmad Ali", "Dayang Siti"],
            "BE2": ["Chong Wei", "Nur Huda"]
        }
        st.warning("⚠️ Note: 'Student_Registry' rows are empty or unreadable. Showing demo registry entries.")

    # 🔐 LIBRARIAN VERIFICATION LAYER
    admin_pass = st.text_input("🗝️ Enter Librarian Credentials to Unlock Tracker Form:", type="password")
    
    if admin_pass == st.secrets["admin_password"]:
        st.success("🔓 Librarian Access Granted.")
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
        reading_dates = []
        cols = st.columns(4)
        
        for i in range(int(article_count)):
            col_index = i % 4
            with cols[col_index]:
                date_val = st.date_input(f"Article #{i+1} Date", key=f"challenge_date_{i}")
                reading_dates.append(date_val.strftime("%d/%m/%Y"))
        
        st.markdown("---")
        
        # Submit Operation
        if st.button("🔥 Submit Tally Data Logs", use_container_width=True):
            try:
                challenge_db = client.open("Articles_Tracker_DB").worksheet("Reading_Article_DB")
                dates_string = ", ".join(reading_dates)
                
                new_tally_row = [
                    datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    selected_class,
                    selected_student,
                    selected_month,
                    int(article_count),
                    dates_string
                ]
                
                # Directly append the row safely without checking for previous elements
                challenge_db.append_row(new_tally_row)
                st.success(f"🎉 Success! Recorded {article_count} articles for {selected_student} ({selected_class}) into the system database.")
                
            except Exception as e:
                st.error(f"Failed to record data. Please check worksheet connection. Details: {e}")
                
    elif admin_pass != "":
        st.error("❌ Invalid Credentials. Access Denied.")
