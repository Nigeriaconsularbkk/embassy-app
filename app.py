import streamlit as st
from docxtpl import DocxTemplate
import io
from datetime import datetime

st.set_page_config(page_title="Embassy Doc System", page_icon="🇳🇬")
st.title("🇳🇬 Embassy Document Generator")

# --- SIDEBAR: CATEGORY SELECTION ---
category = st.sidebar.selectbox("Select Category", ["Visa", "Land Transport", "Visa Transfer"])

context = {}
template_file = ""
today_date = datetime.now().strftime("%d %B %Y")

# --- SHARED INPUTS (Common to most docs) ---
name = st.text_input("Full Name")
passport = st.text_input("Passport Number")
dob = st.text_input("Date of Birth (e.g., 01 Jan 1990)")
pob = st.text_input("Place of Birth (pob)")

# --- GENDER LOGIC ---
gender_choice = st.radio("Gender", ["Male", "Female"], horizontal=True)
if gender_choice == "Male":
    g1, g2, g3 = "he", "his", "him"
else:
    g1, g2, g3 = "she", "her", "her"

# --- CATEGORY LOGIC ---
if category == "Visa":
    sub_visa = st.selectbox("Visa Type", ["30 Days Extension", "Student", "Employment", "Marriage"])
    
    if sub_visa == "30 Days Extension":
        template_file = "visa_30days.docx"
        leave_on = st.text_input("Leave on (Date)")
        context = {"name": name, "passport": passport, "dob": dob, "pob": pob, "leave_on": leave_on}
        
    elif sub_visa == "Student":
        template_file = "visa_student.docx"
        program = st.text_input("Program (e.g., Bachelor of Arts)")
        place = st.text_input("Place of Study")
        location = st.text_input("Location of Study")
        context = {"name": name, "passport": passport, "dob": dob, "pob": pob, "program": program, "place_of_study": place, "location_of_study": location}

    elif sub_visa == "Employment":
        template_file = "visa_employment.docx"
        place_work = st.text_input("Place of Work")
        loc_work = st.text_input("Location of Work")
        p_issue = st.text_input("Passport Place of Issue")
        c_issue = st.text_input("Country of Issue")
        d_issue = st.text_input("Date of Issue")
        p_expiry = st.text_input("Passport Expiration")
        context = {"name": name, "dob": dob, "pob": pob, "passport": passport, "place_of_issue": p_issue, "country_of_issue": c_issue, "date_of_issue": d_issue, "passport_expiration": p_expiry, "place_of_work": place_work, "location_of_work": loc_work}

    elif sub_visa == "Marriage":
        template_file = "visa_marriage.docx"
        context = {"name": name, "passport": passport, "dob": dob, "pob": pob}

elif category == "Land Transport":
    template_file = "land_transport.docx"
    land_purpose = st.selectbox("Select Purpose", ["transferring a vehicle as requested", "registering a driving license as requested"])
    address = st.text_area("Current Address")
    context = {"name": name, "passport": passport, "dob": dob, "pob": pob, "current_address": address, "purpose": land_purpose}

elif category == "Visa Transfer":
    template_file = "visa_transfer.docx"
    p_issue = st.text_input("Passport Place of Issue")
    d_issue = st.text_input("Date of Issue")
    p_expiry = st.text_input("Passport Expiration")
    old_pp = st.text_input("Old Passport Number")
    old_pp_exp = st.text_input("Old Passport Expiration Date")
    context = {"name": name, "dob": dob, "pob": pob, "passport": passport, "place_of_issue": p_issue, "date_of_issue": d_issue, "passport_expiration": p_expiry, "old_passport": old_pp, "old_passport_expiration": old_pp_exp}

# Add Universal Tags (Gender and Date) to all contexts
context.update({"gender1": g1, "gender2": g2, "gender3": g3, "date": today_date})

# --- GENERATION BUTTON ---
st.divider()
if st.button("🚀 Generate Document"):
    if not name or not passport:
        st.error("Missing essential info: Name and Passport are required.")
    else:
        try:
            doc = DocxTemplate(template_file)
            doc.render(context)
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            st.balloons()
            st.success(f"Document for {name} is ready!")
            st.download_button(
                label="Download Word Doc 📥",
                data=bio,
                file_name=f"{template_file.split('.')[0]}_{name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error: Make sure '{template_file}' is uploaded to GitHub. (Details: {e})")
