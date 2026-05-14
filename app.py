import streamlit as st
from docxtpl import DocxTemplate
import io
import zipfile
from datetime import datetime
from docx import Document  # Ensure you have 'python-docx' in requirements.txt

# 1. PAGE SETUP
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
st.markdown("<h1 class='main-title'>EMBASSY OF NIGERIA BKK SYSTEM</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: #555;'>Consular & Immigration Department</h4>", unsafe_allow_html=True)

# --- TABS NAVIGATION ---
tab1, tab2 = st.tabs(["📝 Individual Generator", "📂 Bulk Prison Updater"])

# ==========================================
# TAB 1: INDIVIDUAL DOCUMENT GENERATOR
# ==========================================
with tab1:
    context = {}
    template_file = ""

    st.subheader("🟢 STEP 1: CHOOSE MAIN CATEGORY")
    category = st.radio(
        "Select the department for this document:",
        ["Visa", "Land Transport", "Visa Transfer"],
        horizontal=True, key="cat_radio"
    )

    st.write("---")
    st.subheader(f"📝 STEP 2: ENTER {category.upper()} DETAILS")

    doc_date = st.date_input("Document Date", value=datetime.now(), key="indiv_date")
    formatted_date = doc_date.strftime("%d %B %Y")

    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Full Name (As shown in Passport)")
        passport = st.text_input("Passport Number")
    with col2:
        dob = st.text_input("Date of Birth (e.g., 01 Jan 1990)")
        pob = st.text_input("Place of Birth (e.g., Lagos, Nigeria)")

    gender_choice = st.radio("Gender of Applicant", ["Male", "Female"], horizontal=True)
    g1, g2, g3 = ("he", "his", "him") if gender_choice == "Male" else ("she", "her", "her")

    if category == "Visa":
        sub_visa = st.selectbox("Purpose of Visa Extension", ["30 Days Extension", "Student", "Employment", "Marriage"])
        if sub_visa == "30 Days Extension":
            template_file = "visa_30days.docx"
            context = {"leave_on": st.text_input("Expected Date of Departure")}
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

    final_context = {
        "name": name,
        "name_capital": name.upper() if name else "",
        "passport": passport,
        "dob": dob,
        "pob": pob,
        "gender1": g1,
        "gender2": g2,
        "gender3": g3,
        "date": formatted_date
    }
    final_context.update(context)

    st.write("---")
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
                st.success(f"Generated successfully for {name}")
                st.download_button(
                    label="📥 Download Word Document",
                    data=bio,
                    file_name=f"{category}_{name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error: {e}")

# ==========================================
# TAB 2: BULK PRISON UPDATER
# ==========================================
with tab2:
    st.subheader("⚡ Bulk Date Swapper for Prison Letters")
    st.info("Upload your old letters. The system will find the old dates and replace them with the new ones across all files.")

    col_old, col_new = st.columns(2)
    with col_old:
        st.markdown("**🔍 OLD CONTENT TO FIND**")
        old_issue = st.text_input("Old Issue Date (exactly as in file)", placeholder="e.g., 12 April 2026")
        old_visit = st.text_input("Old Visit Month/Date", placeholder="e.g., April 2026")
    
    with col_new:
        st.markdown("**🖋️ NEW CONTENT TO REPLACE**")
        new_issue = st.text_input("New Issue Date", value=datetime.now().strftime("%d %B %Y"))
        new_visit = st.text_input("New Visit Month/Date", placeholder="e.g., May 2026")

    uploaded_prison_files = st.file_uploader("Upload all 50+ Prison Files", type="docx", accept_multiple_files=True)

    if st.button("🚀 UPDATE ALL PRISON LETTERS"):
        if not uploaded_prison_files or not old_issue or not new_issue:
            st.warning("Please fill in the 'Find' and 'Replace' dates and upload files.")
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for uploaded_file in uploaded_prison_files:
                    doc = Document(uploaded_file)
                    
                    # Search and replace in paragraphs and tables
                    for section in [doc.paragraphs] + [p for table in doc.tables for row in table.rows for cell in row.cells for p in cell.paragraphs]:
                        for p in section if isinstance(section, list) else [section]:
                            if old_issue in p.text:
                                p.text = p.text.replace(old_issue, new_issue)
                            if old_visit and old_visit in p.text:
                                p.text = p.text.replace(old_visit, new_visit)
                    
                    out_stream = io.BytesIO()
                    doc.save(out_stream)
                    zip_file.writestr(uploaded_file.name, out_stream.getvalue())

            st.success(f"✅ Successfully processed {len(uploaded_prison_files)} files!")
            st.download_button(
                label="📥 Download Updated Letters (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"Updated_Prison_Letters_{new_visit}.zip",
                mime="application/zip"
            )
