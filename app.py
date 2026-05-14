import streamlit as st
from docxtpl import DocxTemplate
import io
import zipfile
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# 1. PAGE SETUP
st.set_page_config(page_title="Embassy of Nigeria BKK", page_icon="🇳🇬", layout="wide")

# --- CUSTOM GREEN THEME CSS ---
st.markdown("""
    <style>
    .main-title { color: #008751; text-align: center; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #008751; padding-bottom: 10px; }
    .stSubheader { color: #008751 !important; font-weight: bold; }
    div.stButton > button:first-child { background-color: #008751; color: white; border-radius: 10px; border: none; height: 3em; width: 100%; font-weight: bold; font-size: 20px; }
    </style>
    """, unsafe_allow_html=True)

# --- FONT HELPER ---
def apply_smart_font(paragraph, text, is_bold=False, is_underline=False):
    chunks = re.findall(r'[a-zA-Z]+|[^a-zA-Z]+', text)
    for chunk in chunks:
        run = paragraph.add_run(chunk)
        run.bold = is_bold
        run.underline = is_underline
        if re.search(r'[a-zA-Z]', chunk):
            run.font.name = 'Times New Roman'
            run.font.size = Pt(13.5)
            r = run._element.rPr
            r.get_or_add_rFonts().set(qn('w:ascii'), 'Times New Roman')
            r.get_or_add_rFonts().set(qn('w:hAnsi'), 'Times New Roman')
        else:
            run.font.name = 'Angsana New'
            run.font.size = Pt(17)
            r = run._element.rPr
            r.get_or_add_rFonts().set(qn('w:cs'), 'Angsana New')
            r.get_or_add_rFonts().set(qn('w:eastAsia'), 'Angsana New')
            r.get_or_add_rFonts().set(qn('w:ascii'), 'Angsana New')
            r.get_or_add_rFonts().set(qn('w:hAnsi'), 'Angsana New')

# --- OFFICIAL HEADER ---
st.markdown("<h1 class='main-title'>EMBASSY OF NIGERIA BKK SYSTEM</h1>", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📝 Individual Generator", "📂 Professional Bulk Updater"])

# ==========================================
# TAB 1: INDIVIDUAL DOCUMENT GENERATOR
# ==========================================
with tab1:
    st.subheader("🟢 STEP 1: SELECT DOCUMENT TYPE")
    category = st.selectbox("Category:", ["Visa 30 Days Extension", "Visa Student", "Visa Employment", "Visa Marriage", "Land Transport", "Visa Transfer"])

    st.write("---")
    
    # Common Fields
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Full Name")
        passport = st.text_input("Passport Number")
        dob = st.text_input("Date of Birth (e.g. 01/01/1990)")
    with col2:
        pob = st.text_input("Place of Birth (City/State)")
        doc_date = st.date_input("Document Issue Date", value=datetime.now())
        gender_choice = st.radio("Gender of Applicant", ["Male", "Female"], horizontal=True)

    # Logic for Pronouns
    # gender1: he/she, gender2: his/her, gender3: him/her
    g1, g2, g3 = ("he", "his", "him") if gender_choice == "Male" else ("she", "her", "her")

    st.write("---")
    st.subheader("🔵 STEP 2: CATEGORY SPECIFIC DETAILS")
    
    # Context Dictionary initialization
    final_context = {
        "name": name, "name_capital": name.upper(), "passport": passport,
        "dob": dob, "pob": pob, "gender1": g1, "gender2": g2, "gender3": g3,
        "date": doc_date.strftime("%d %B %Y")
    }

    template_file = ""

    # Dynamic Inputs based on Template
    if category == "Visa 30 Days Extension":
        template_file = "visa_30days.docx"
        final_context["leave_on"] = st.text_input("Intended Leave Date (Thailand)")

    elif category == "Visa Student":
        template_file = "visa_student.docx"
        final_context["program"] = st.text_input("Program of Study (e.g., MBA)")
        final_context["place_of_study"] = st.text_input("University/School Name")
        final_context["location_of_study"] = st.text_input("City/Province of School")

    elif category == "Visa Employment":
        template_file = "visa_employment.docx"
        final_context["place_of_work"] = st.text_input("Company/Employer Name")
        final_context["location_of_work"] = st.text_input("Job Location (e.g. Bangkok)")
        final_context["place_of_issue"] = st.text_input("Passport Place of Issue")
        final_context["country_of_issue"] = st.text_input("Passport Country of Issue", value="Nigeria")
        final_context["date_of_issue"] = st.text_input("Passport Date of Issue")
        final_context["passport_expiration"] = st.text_input("Passport Expiry Date")

    elif category == "Visa Marriage":
        template_file = "visa_marriage.docx"
        # Uses standard common fields + gender tags

    elif category == "Land Transport":
        template_file = "land_transport.docx"
        final_context["current_address"] = st.text_area("Current Residential Address in Thailand")
        final_context["purpose"] = st.text_input("Specific Assistance Required (e.g., obtaining a driving license)")

    elif category == "Visa Transfer":
        template_file = "visa_transfer.docx"
        final_context["old_passport"] = st.text_input("Previous Passport Number")
        final_context["old_passport_expiration"] = st.text_input("Previous Passport Expiry Date")
        final_context["place_of_issue"] = st.text_input("New Passport Place of Issue")
        final_context["date_of_issue"] = st.text_input("New Passport Date of Issue")
        final_context["passport_expiration"] = st.text_input("New Passport Expiry Date")

    st.write("---")
    if st.button("💾 GENERATE DOCUMENT"):
        if not name or not passport:
            st.warning("Please enter at least Name and Passport Number.")
        else:
            try:
                doc = DocxTemplate(template_file)
                doc.render(final_context)
                bio = io.BytesIO()
                doc.save(bio)
                st.success(f"Successfully generated {template_file} for {name}")
                st.download_button("📥 Download Document", bio.getvalue(), f"{category.replace(' ', '_')}_{name}.docx")
            except Exception as e:
                st.error(f"Error: {e}. Check if {template_file} exists in your directory.")

# ==========================================
# TAB 2: PROFESSIONAL BULK UPDATER (PRISON)
# ==========================================
with tab2:
    st.subheader("🏛️ Smart Prison Batch Updater")
    st.info("Use this to update multiple Thai-language prison visit letters at once.")
    
    col_a, col_b = st.columns(2)
    with col_a:
        ref_id = st.text_input("Search for Reference line:", value="อีเอ็นบี/ซีเอ็น.07")
        new_issue = st.text_input("New Date of Issue (Thai/Eng)", value=datetime.now().strftime("%d %B %Y"))
    
    with col_b:
        new_visit = st.text_input("New Month & Time Range", placeholder="มิถุนายน 2569 เวลา 12.00 - 15.00 น.")

    files = st.file_uploader("Upload Word Docs", type=["docx"], accept_multiple_files=True)

    if st.button("🚀 BATCH UPDATE ALL"):
        if not files or not new_visit:
            st.error("Please upload files and provide visit details.")
        else:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zip_f:
                for f in files:
                    doc = Document(f)
                    update_next = False
                    
                    for p in list(doc.paragraphs):
                        if ref_id in p.text:
                            update_next = True
                            continue
                        if update_next and len(p.text.strip()) > 0:
                            p.text = ""
                            apply_smart_font(p, new_issue)
                            update_next = False
                            continue
                        
                        if "ในเดือน" in p.text:
                            parts = p.text.split("ในเดือน", 1)
                            p.text = ""
                            apply_smart_font(p, parts[0] + "ในเดือน")
                            apply_smart_font(p, " ")
                            apply_smart_font(p, new_visit, is_bold=True, is_underline=True)

                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    if "ในเดือน" in p.text:
                                        parts = p.text.split("ในเดือน", 1)
                                        p.text = ""
                                        apply_smart_font(p, parts[0] + "ในเดือน")
                                        apply_smart_font(p, " ")
                                        apply_smart_font(p, new_visit, is_bold=True, is_underline=True)

                    out = io.BytesIO()
                    doc.save(out)
                    zip_f.writestr(f.name, out.getvalue())
            
            st.success("Batch Complete!")
            st.download_button("📥 Download Updated ZIP", zip_buf.getvalue(), "Prison_Letters_Updated.zip")
