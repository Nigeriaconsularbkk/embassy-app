import streamlit as st
from docxtpl import DocxTemplate
import io
import os
from datetime import datetime

# Page configuration
st.set_page_config(page_title="Embassy Letter System", page_icon="🇳🇬")
st.title("🇳🇬 Embassy Visa Extension System")

TEMPLATE_NAME = "template.docx"

if not os.path.exists(TEMPLATE_NAME):
    st.error(f"❌ Cannot find '{TEMPLATE_NAME}'! Please make sure it is in the LetterProject folder.")
else:
    with st.form("visa_form"):
        st.subheader("Applicant Details")
        name = st.text_input("Full Name")
        
        col1, col2 = st.columns(2)
        with col1:
            passport = st.text_input("Passport Number")
            dob = st.text_input("Date of Birth (e.g. 01/01/1990)")
            gender_input = st.radio("Select Gender", ["Male", "Female"])
        
        with col2:
            pob = st.text_input("Place of Birth (e.g. Lagos)")
            visa_until = st.text_input("Visa Expiry Date (e.g. 20 May 2026)")
            letter_date = st.date_input("Letter Date", value=datetime.now())

        submit = st.form_submit_button("🚀 Generate Official Letter")

    if submit:
        if not name or not passport:
            st.warning("⚠️ Please enter at least the Name and Passport Number.")
        else:
            try:
                doc = DocxTemplate(TEMPLATE_NAME)

                # GENDER LOGIC
                if gender_input == "Male":
                    g1 = "his"
                    g2 = "him"
                else:
                    g1 = "her"
                    g2 = "her"

                # Dictionary matching your Word Template tags
                context = {
                    "date": letter_date.strftime("%d %B %Y"),
                    "name": name,
                    "passport": passport,
                    "dob": dob,
                    "pob": pob,
                    "gender1": g1,      # Fills {{gender1}}
                    "visauntil": visa_until,
                    "gender2": g2       # Fills {{gender2}}
                }
                
                doc.render(context)
                
                bio = io.BytesIO()
                doc.save(bio)
                bio.seek(0)
                
                st.balloons()
                st.success(f"✨ Official Letter for {name} generated!")
                
                st.download_button(
                    label="Download Word Doc 📥",
                    data=bio,
                    file_name=f"Visa_Extension_{name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error: {e}")