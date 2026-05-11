import streamlit as st
from docxtpl import DocxTemplate
import io
from datetime import datetime

# 1. PAGE SETUP (Must be the first Streamlit command)
st.set_page_config(page_title="Embassy of Nigeria BKK", page_icon="🇳🇬", layout="wide")

# --- CUSTOM GREEN THEME CSS ---
st.markdown("""
    <style>
    .main-title {
        color: #008751;
        text-align: center;
        font-family: 'Arial Black', sans-serif;
        border-bottom: 3px solid #008751;
        padding-bottom: 10px;
    }
    .stSubheader {
        color: #008751 !important;
        font-weight: bold;
    }
    div.stButton > button:first-child {
        background-color: #008751;
        color: white;
        border-radius: 10px;
        border: none;
        height: 3em;
        width: 100%;
        font-weight: bold;
        font-size: 20px;
    }
    div.stButton > button:hover {
        background-color: #005c37;
        color: white;
        border: 1px solid #008751;
    }
    </style>
    """, unsafe_allow_html=True)

# --- OFFICIAL HEADER ---
st.markdown("<h1 class='main-title'>EMBASSY OF NIGERIA BKK DOCUMENT GENERATING SYSTEM</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: #555;'>Consular & Immigration Department</h4>", unsafe_allow_html=True)

# --- INITIALIZE VARIABLES ---
context = {}
template_file = ""
today_date = datetime.now().strftime("%d %B %Y")

# --- STEP 1: CATEGORY SELECTION ---
st.write("---")
st.subheader("🟢 STEP 1: CHOOSE MAIN CATEGORY")
category = st.radio(
    "Select the department for this document:",
    ["Visa", "Land Transport", "Visa Transfer"],
    horizontal=True
)

# --- STEP 2: FORM INPUTS ---
st.write("---")
st.subheader(f"📝 STEP 2: ENTER {category.upper()} DETAILS")

col1, col2 = st.columns(2)
with col1:
    name = st.text_input("Full Name (As shown in Passport)")
    passport = st.text_input("Passport Number")
with col2:
    dob = st.text_input("Date of Birth (e.g., 01 Jan 1990)")
    pob = st.text_input("Place of Birth (e.g., Lagos, Nigeria)")

gender_choice = st.radio("Gender of Applicant", ["Male", "Female"], horizontal=True)
g1, g2, g3 = ("he", "his", "him") if gender_choice == "Male" else ("she", "her", "her")

# --- CATEGORY SPECIFIC LOGIC ---
if category == "Visa":
    sub_visa = st.selectbox("Purpose of Visa Extension", ["30 Days Extension", "Student", "Employment", "Marriage"])
    
    if sub_visa == "30 Days Extension":
        template_file = "visa_30days.docx"
        leave_on = st.text_input("Expected Date of Departure")
        context = {"leave_on": leave_on}
        
    elif sub_visa == "Student":
        template_file = "visa_student.docx"
        context = {
            "program": st.text_input("Program of Study"), 
            "place_of_study": st.text_input("Name of University/School"), 
            "location_of_study": st.text_input("Location of School")
        }

    elif sub_visa == "Employment":
        template_file = "visa_employment.docx"
        context = {
            "place_of_work": st.text_input("Place of Work (Company Name)"), 
            "location_of_work": st.text_input("Work Location"), 
            "place_of_issue": st.text_input("Passport Place of Issue"), 
            "country_of_issue": st.text_input("Country of Issue"), 
            "date_of_issue": st.text_input("Date of Passport Issue"), 
            "passport_expiration": st.text_input("Passport Expiration Date")
        }

    elif sub_visa == "Marriage":
        template_file = "visa_marriage.docx"

elif category == "Land Transport":
    template_file = "land_transport.docx"
    context = {
        "purpose": st.selectbox("Action Requested", ["transferring a vehicle as requested", "registering a driving license as requested"]),
        "current_address": st.text_area("Resident Address in Thailand")
    }

elif category == "Visa Transfer":
    template_file = "visa_transfer.docx"
    context = {
        "place_of_issue": st.text_input("Place of Issue"), 
        "date_of_issue": st.text_input("Date of Issue (New PP)"), 
        "passport_expiration": st.text_input("Expiration Date (New PP)"), 
        "old_passport": st.text_input("Old Passport Number"), 
        "old_passport_expiration": st.text_input("Old Passport Expiration Date")
    }

# --- MERGE ALL CONTEXT ---
final_context = {
    "name": name,
    "name_capital": name.upper() if name else "",
    "passport": passport,
    "dob": dob,
    "pob": pob,
    "gender": g1,
    "gender1": g1,
    "gender2": g2,
    "gender3": g3,
    "date": today_date
}
final_context.update(context)

# --- STEP 3: GENERATE ---
st.write("---")
st.subheader("📂 STEP 3: GENERATE DOCUMENT")

if st.button("💾 CLICK TO GENERATE & DOWNLOAD"):
    if not name.strip() or not passport.strip():
        st.error("⚠️ Mandatory fields missing: Full Name and Passport Number.")
    else:
        try:
            doc = DocxTemplate(template_file)
            doc.render(final_context)
            
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            st.balloons()
            st.success(f"Generated for {name}")
            st.download_button(
                label="📥 Download Word Document",
                data=bio,
                file_name=f"{category}_{name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error: {e}")
