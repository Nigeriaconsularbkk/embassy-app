import streamlit as st
from docxtpl import DocxTemplate
import io
import zipfile
from datetime import datetime
from docx import Document

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
        horizontal=True, key="indiv_cat"
    )

    st.write("---")
    st.subheader(f"📝 STEP 2: ENTER {category.upper()} DETAILS")

    doc_date = st.date_input("Document Date", value=datetime.now(), key="indiv_date_picker")
    formatted_date = doc_date.strftime("%d %B %Y")

    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Full Name (As shown in Passport)", key="indiv_name")
        passport = st.text_input("Passport Number", key="indiv_pp")
    with col2:
        dob = st.text_input("Date of Birth (e.g., 01 Jan 1990)", key="indiv_dob")
        pob = st.text_input("Place of Birth (e.g., Lagos, Nigeria)", key="indiv_pob")

    gender_choice = st.radio("Gender of Applicant", ["Male", "Female"], horizontal=True, key="indiv_gender")
    g1, g2, g3 = ("he", "his", "him") if gender_choice == "Male" else ("she", "her", "her")

    if category == "Visa":
        sub_visa = st.selectbox("Purpose", ["30 Days Extension", "Student", "Employment", "Marriage"])
        if sub_visa == "30 Days Extension":
            template_file = "visa_30days.docx"
            context = {"leave_on": st.text_input("Expected Departure Date")}
        elif sub_visa == "Student":
            template_file = "visa_student.docx"
            context = {
                "program": st.text_input("Program of Study"), 
                "place_of_study": st.text_input("School Name"), 
                "location_of_study": st.text_input("School Location")
            }
        elif sub_visa == "Employment":
            template_file = "visa_employment.docx"
            context = {
                "place_of_work": st.text_input("Company Name"), 
                "location_of_work": st.text_input("Work Location"), 
                "place_of_issue": st.text_input("Passport Place of Issue"), 
                "country_of_issue": st.text_input("Country of Issue"), 
                "date_of_issue": st.text_input("Date of Issue"), 
                "passport_expiration": st.text_input("Expiration Date")
            }
        elif sub_visa == "Marriage":
            template_file = "visa_marriage.docx"

    elif category == "Land Transport":
        template_file = "land_transport.docx"
        context = {
            "purpose": st.selectbox("Action", ["transferring a vehicle as requested", "registering a driving license as requested"]),
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
        "name": name, "name_capital": name.upper() if name else "",
        "passport": passport, "dob": dob, "pob": pob,
        "gender1": g1, "gender2": g2, "gender3": g3, "date": formatted_date
    }
    final_context.update(context)

    if st.button("💾 GENERATE SINGLE DOCUMENT"):
        if not name.strip() or not passport.strip():
            st.error("Missing Full Name or Passport.")
        else:
            try:
                doc_templ = DocxTemplate(template_file)
                doc_templ.render(final_context)
                bio = io.BytesIO()
                doc_templ.save(bio)
                st.balloons()
                st.download_button("📥 Download Document", bio.getvalue(), f"{name}.docx")
            except Exception as e:
                st.error(f"Error: {e}")

# ==========================================
# TAB 2: BULK PRISON UPDATER (Keeps Angsana)
# ==========================================
with tab2:
    st.subheader("⚡ Bulk Date Swapper for Prison Letters")
    st.info("Replaces dates while preserving your original **Angsana New** font and size.")

    col_old, col_new = st.columns(2)
    with col_old:
        st.markdown("**🔍 FIND**")
        old_issue = st.text_input("Old Issue Date", placeholder="e.g., 10 May 2026", key="old_i")
        old_visit = st.text_input("Old Visit Date/Month", placeholder="e.g., May 2026", key="old_v")
    
    with col_new:
        st.markdown("**🖋️ REPLACE**")
        new_issue = st.text_input("New Issue Date", value=datetime.now().strftime("%d %B %Y"), key="new_i")
        new_visit = st.text_input("New Visit Date/Month", placeholder="e.g., June 2026", key="new_v")

    uploaded_prison_files = st.file_uploader("Upload all .docx files", type=["docx"], accept_multiple_files=True)

    if st.button("🚀 BATCH UPDATE ALL LETTERS"):
        if not uploaded_prison_files or not old_issue or not new_issue:
            st.warning("Please fill all date fields and upload files.")
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for uploaded_file in uploaded_prison_files:
                    doc = Document(uploaded_file)
                    
                    # Function to replace text while preserving Run formatting (Font/Size)
                    def safe_replace(paragraphs, old_text, new_text):
                        for p in paragraphs:
                            if old_text in p.text:
                                for run in p.runs:
                                    if old_text in run.text:
                                        run.text = run.text.replace(old_text, new_text)

                    # Replace in Paragraphs
                    safe_replace(doc.paragraphs, old_issue, new_issue)
                    if old_visit: safe_replace(doc.paragraphs, old_visit, new_visit)
                    
                    # Replace in Tables
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                safe_replace(cell.paragraphs, old_issue, new_issue)
                                if old_visit: safe_replace(cell.paragraphs, old_visit, new_visit)
                    
                    out_stream = io.BytesIO()
                    doc.save(out_stream)
                    zip_file.writestr(uploaded_file.name, out_stream.getvalue())

            st.success(f"Batch update complete for {len(uploaded_prison_files)} files!")
            st.download_button("📥 Download Updated ZIP", zip_buffer.getvalue(), f"Batch_Update_{new_visit}.zip")
