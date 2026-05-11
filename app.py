import streamlit as st
from docxtpl import DocxTemplate
import io
from datetime import datetime

# Page configuration
st.set_page_config(page_title="Embassy of Nigeria BKK", page_icon="🇳🇬", layout="wide")

# --- CUSTOM GREEN THEME CSS ---
st.markdown("""
    <style>
    /* Main title color */
    .main-title {
        color: #008751;
        text-align: center;
        font-family: 'Arial Black', sans-serif;
        border-bottom: 3px solid #008751;
        padding-bottom: 10px;
    }
    /* Subheader color */
    .stSubheader {
        color: #008751 !important;
        font-weight: bold;
    }
    /* Button styling */
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
    /* Radio button selection color */
    div[data-baseweb="radio"] > div {
        color: #008751 !important;
    }
    </style>
    """, unsafe_content_type=True)

# --- OFFICIAL HEADER ---
st.markdown("<h1 class='main-title'>EMBASSY OF NIGERIA BKK DOCUMENT GENERATING SYSTEM</h1>", unsafe_content_type=True)
st.markdown("<h4 style='text-align: center; color: #555;'>Consular & Immigration Department</h4>", unsafe_content_type=True)
st.write("")

# --- STEP 1: CATEGORY SELECTION ---
st.subheader("🟢 STEP 1: CHOOSE MAIN CATEGORY")
category = st.radio(
    "Select the department for this document:",
    ["Visa", "Land Transport", "Visa Transfer"],
    horizontal=True
)

st.write("---")

# --- INITIALIZE VARIABLES ---
context = {}
template_file = ""
today_date = datetime.now().strftime("%d %B %Y")

# --- STEP 2: FORM INPUTS ---
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
        program = st.text_input("Program of Study")
        place = st.text_input("Name of University/School")
        loc = st.text_input("Location of School")
        context = {"program": program, "place_of_study": place, "location_of_study": loc}

    elif sub_visa == "Employment":
        template_file = "visa_employment.docx"
        p_work = st.text_input("Place of Work (Company Name)")
        l_work = st.text_input("Work Location")
        p_issue = st.text_input("Passport Place of Issue")
        c_issue = st.text_input("Country of Issue")
        d_issue = st.text_input("Date of Passport Issue")
        p_expiry = st.text_input("Passport Expiration Date")
        context = {
            "place_of_issue": p_issue, "country_of_issue": c_issue, 
            "date_of_issue": d_issue, "passport_expiration": p_expiry, 
            "place_of_work": p_work, "location_of_work": l_work
        }

    elif sub_visa == "Marriage":
        template_file = "visa_marriage.docx"
        context = {}

elif category == "Land Transport":
    template_file = "land_transport.docx"
    land_purpose = st.selectbox("Action Requested", ["transferring a vehicle as requested", "registering a driving license as requested"])
    address = st.text_area("Resident Address in Thailand")
    context = {"current_address": address, "purpose": land_purpose}

elif category == "Visa Transfer":
    template_file = "visa_transfer.docx"
    p_issue = st.text_input("Place of Issue")
    d_issue = st.text_input("Date of Issue (New Passport)")
    p_expiry = st.text_input("Expiration Date (New Passport)")
    old_pp = st.text_input("Old Passport Number")
    old_pp_exp = st.text_input("Old Passport Expiration Date")
    context = {
        "place_of_issue": p_issue, "date_of_issue": d_issue, 
        "passport_expiration": p_expiry, "old_passport": old_pp, 
        "old_passport_expiration": old_pp_exp
    }

# Combine data
final_context = {
    "name": name, "passport": passport, "dob": dob, "pob": pob,
    "gender1": g1, "gender2": g2, "gender3": g3, "date": today_date
}
final_context.update(context)

# --- STEP 3: FINAL GENERATION ---
st.write("---")
st.subheader("📂 STEP 3: GENERATE DOCUMENT")
if st.button("💾 CLICK TO GENERATE & DOWNLOAD"):
    if not name or not passport:
        st.error("Missing Information: Name and Passport Number are mandatory.")
    else:
        try:
            doc = DocxTemplate(template_file)
            doc.render(final_context)
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            st.balloons()
            st.success(f"Successfully generated document for {name}")
            st.download_button(
                label=f"📥 Download {category} Letter",
                data=bio,
                file_name=f"{category.replace(' ', '_')}_{name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error: {e}")
