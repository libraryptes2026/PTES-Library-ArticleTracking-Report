import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

# --- 1. PAGE CONFIGURATION & STYLING ---
st.set_page_config(
    page_title="PTES Library Services Portal",
    page_icon="📚",
    layout="wide"
)

# --- 2. SECURE DATABASE CONNECTION (GOOGLE SHEETS) ---
def connect_to_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        # Pull credentials safely from the top-most global settings of Streamlit Secrets
        creds_info = st.secrets["gspread_creds"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Database Connection Error: {e}")
        return None

client = connect_to_sheets()

# --- 3. SIDEBAR BRANDING & POLICIES ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c3/Flag_of_Brunei.svg/180px-Flag_of_Brunei.svg.png", width=100)
    st.title("PTES Services")
    st.markdown("### 📋 User Guidelines")
    
    with st.expander("🛡️ Room Reservation Rules"):
        st.write("""
        1. **Lecturer Presence:** A lecturer must be present at all times.
        2. **Key Management:** Collect and return keys to the library counter.
        3. **Cleanliness:** Clear whiteboards and leave chairs arranged.
        """)
        
    with st.expander("📖 Reading Challenge Rules"):
        st.write("""
        1. **Honesty First:** Tallies must correspond strictly to verified manuscript reads.
        2. **Librarian Override:** Only authorized librarians can submit tallies.
        """)
    
    # --- PTES FOOTER IN SIDEBAR ---
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

# --- 4. MAIN INTERFACE FRAMEWORK ---
st.title("🏫 PTES Digital Services Hub")
st.markdown("---")

# Split the application into two dedicated feature tabs
tab1, tab2 = st.tabs(["📅 Room Reservations", "📚 Reading Challenge Tally"])

# ==========================================
# FEATURE TAB 1: DISCUSSION ROOM RESERVATIONS
# ==========================================
with tab1:
    st.header("Discussion Room Booking Engine")
    
    if client:
        try:
            # Connect to Room Reservation worksheet
            booking_sheet = client.open("Library_Booking_DB").sheet1
            data = booking_sheet.get_all_records()
            df = pd.DataFrame(data)
            
            # Form UI for Room Bookings
            col1, col2 = st.columns(2)
            with col1:
                room_selection = st.selectbox("Select Discussion Room:", ["Room A (Max 6 pax)", "Room B (Max 10 pax)"])
                booking_date = st.date_input("Select Date:", min_value=datetime.today())
            with col2:
                time_slot = st.selectbox("Select Time Slot:", ["08:00 - 09:00", "09:00 - 10:00", "10:00 - 11:00", "11:00 - 12:00", "13:00 - 14:00", "14:00 - 15:00"])
                lecturer_name = st.text_input("Lecturer In Charge:")
                
            if st.button("Confirm Room Booking"):
                if lecturer_name.strip() == "":
                    st.warning("Please fill in the Lecturer In Charge name.")
                else:
                    # Check for existing double bookings
                    is_clash = False
                    if not df.empty:
                        clash_check = df[(df['Room'] == room_selection) & (df['Date'] == booking_date.strftime("%d/%m/%Y")) & (df['Time Slot'] == time_slot)]
                        if not clash_check.empty:
                            is_clash = True
                            
                    if is_clash:
                        st.error(f"❌ Structural Clash! {room_selection} is already reserved for {time_slot} on this day.")
                    else:
                        new_booking = [datetime.now().strftime("%d/%m/%Y %H:%M:%S"), room_selection, booking_date.strftime("%d/%m/%Y"), time_slot, lecturer_name]
                        booking_sheet.append_row(new_booking)
                        st.success(f"🎉 Reservation Successful for {room_selection}!")
                        st.rerun()
                        
            # Live Reservation Board Presentation
            st.markdown("### 📊 Live Reservation Board")
            if not df.empty:
                st.dataframe(df[["Room", "Date", "Time Slot", "Lecturer In Charge"]], use_container_width=True)
            else:
                st.info("No current reservations found.")
                
        except Exception as e:
            st.error(f"Reservation Engine Error: {e}")

# ==========================================
# FEATURE TAB 2: READING CHALLENGE DASHBOARD
# ==========================================
with tab2:
    st.header("PTES Annual Reading Challenge")
    
    if client:
        # Dynamic Student Registry Aggregator
        try:
            registry_sheet = client.open("Library_Booking_DB").worksheet("Student_Registry")
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
            # Secure offline backup loop if worksheet hasn't been instantiated yet
            class_registry = {
                "BE1": ["Ahmad Ali", "Dayang Siti"],
                "BE2": ["Chong Wei", "Nur Huda"]
            }
            st.warning("⚠️ Note: 'Student_Registry' worksheet not found in Google Sheets. Running offline sample profile.")

        # Admin Verification Layer
        admin_pass = st.text_input("🗝️ Enter Librarian Credentials to Unlock:", type="password")
        
        # Pull protected admin password securely from secrets vault
        if admin_pass == st.secrets["admin_password"]:
            st.success("🔓 Librarian Authentication Granted.")
            st.markdown("---")
            
            # Input Fields
            sorted_classes = sorted(list(class_registry.keys()))
            selected_class = st.selectbox("1. Select Form Class:", sorted_classes, key="challenge_class")
            
            student_options = sorted(class_registry.get(selected_class, []))
            selected_student = st.selectbox("2. Select Student Name:", student_options, key="challenge_student")
            
            months_list = ["March", "April", "May", "June", "July", "August", "September", "October"]
            selected_month = st.selectbox("3. Select Challenge Month:", months_list)
            
            article_count = st.number_input("4. Amount of Reading Articles Done:", min_value=1, max_value=31, step=1)
            
            st.write(f"5. Enter the specific calendar dates for the {int(article_count)} materials read:")
            reading_dates = []
            cols = st.columns(3)
            
            for i in range(int(article_count)):
                col_index = i % 3
                with cols[col_index]:
                    date_val = st.date_input(f"Article #{i+1} Date", key=f"challenge_date_{i}")
                    reading_dates.append(date_val.strftime("%d/%m/%Y"))
            
            if st.button("🔥 Submit Tally & Update Leaderboard"):
                try:
                    challenge_db = client.open("Library_Booking_DB").worksheet("Reading_Challenge_DB")
                    dates_string = ", ".join(reading_dates)
                    
                    new_tally_row = [
                        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                        selected_class,
                        selected_student,
                        selected_month,
                        int(article_count),
                        dates_string
                    ]
                    challenge_db.append_row(new_tally_row)
                    st.success(f"🎉 Transaction Recorded! Appended {article_count} entries to {selected_student}'s profile.")
                except Exception as e:
                    st.error(f"Failed to submit entry. Please ensure you created the 'Reading_Challenge_DB' worksheet tab. Details: {e}")
        
        elif admin_pass != "":
            st.error("❌ Incorrect Librarian Password. Access Denied.")
