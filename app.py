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
tab1, tab2 = st.tabs(["📝 Individual Generator", "📂 Smart Bulk Prison Updater"])

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
        horizontal=True, key="indiv_cat_choice"
    )

    st.write("---")
    st.subheader(f"📝 STEP 2: ENTER {category.upper()} DETAILS")

    doc_date = st.date_input("Document Date", value=datetime.now(), key="indiv_date_picker")
    formatted_date = doc_date.strftime("%d %B %Y")

    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Full Name (As shown in Passport)", key="name_input")
        passport = st.text_input("Passport Number", key="pass_input")
    with col2:
        dob = st.text_input("Date of Birth (e.g., 01 Jan 1990)", key="dob_input")
        pob = st.text_input("Place of Birth (e.g., Lagos, Nigeria)", key="pob_input")

    gender_choice = st.radio("Gender of Applicant", ["Male", "Female"], horizontal=True, key="gender_input")
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
                "place_of_study": st.text_input("Name of School"), 
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
        "name": name, "name_capital": name.upper() if name else "",
        "passport": passport, "dob": dob, "pob": pob,
        "gender1": g1, "gender2": g2, "gender3": g3, "date": formatted_date
    }
    final_context.update(context)

    if st.button("💾 GENERATE DOCUMENT", key="gen_btn"):
        if not name.strip() or not passport.strip():
            st.error("Missing mandatory fields.")
        else:
            try:
                doc_templ = DocxTemplate(template_file)
                doc_templ.render(final_context)
                bio = io.BytesIO()
                doc_templ.save(bio)
                st.balloons()
                st.download_button("📥 Download", bio.getvalue(), f"{name}.docx")
            except Exception as e:
                st.error(f"Error: {e}")

# ==========================================
# TAB 2: SMART BULK UPDATER (REGEX + ANGSANA)
# ==========================================
with tab2:
    st.subheader("⚡ Smart Thai Date & Time Updater")
    st.info("💡 **Smart Logic:** This tool finds the word **'ในเดือน'** and replaces everything after it in that line. No need to know the old date!")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**🔍 ISSUE DATE (Top of Letter)**")
        old_issue_date = st.text_input("Find Old Issue Date", placeholder="เช่น 29 พฤษภาคม 2569", key="find_issue")
        new_issue_date = st.text_input("Replace with New Issue Date", value=datetime.now().strftime("%d %B %Y"), key="replace_issue")
    
    with col_b:
        st.markdown("**🖋️ VISIT DETAILS (After 'ในเดือน')**")
        new_visit_details = st.text_input("New Visit Month & Time", placeholder="เช่น มิถุนายน 2569 เวลา 12.00 - 15.00 น.", key="replace_visit")

    uploaded_prison_files = st.file_uploader("Upload Thai .docx files", type=["docx"], accept_multiple_files=True, key="bulk_upload")

    if st.button("🚀 EXECUTE SMART BATCH UPDATE", key="bulk_btn"):
        if not uploaded_prison_files or not new_visit_details:
            st.warning("Please upload files and enter the new visit details.")
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_f:
                for f in uploaded_prison_files:
                    doc = Document(f)
                    
                    def apply_thai_style(paragraph, updated_text):
                        """Clears paragraph and applies Angsana New 16pt with Thai script support."""
                        paragraph.text = ""
                        run = paragraph.add_run(updated_text)
                        run.font.name = 'Angsana New'
                        run.font.size = Pt(16)
                        r = run._element.rPr
                        # Apply fonts to all script types (Complex, East Asia, etc.)
                        for font_type in ['w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs']:
                            r.get_or_add_rFonts().set(qn(font_type), 'Angsana New')

                    def process_paragraphs(paragraphs):
                        for p in paragraphs:
                            # 1. Replace the Issue Date
                            if old_issue_date and old_issue_date in p.text:
                                text_out = p.text.replace(old_issue_date, new_issue_date)
                                apply_thai_style(p, text_out)

                            # 2. Smart Replace after 'ในเดือน'
                            if "ในเดือน" in p.text:
                                # Regex looks for "ในเดือน" plus any text following it
                                pattern = r"(ในเดือน\s*).*"
                                replacement = r"\1" + new_visit_details
                                text_out = re.sub(pattern, replacement, p.text)
                                apply_thai_style(p, text_out)

                    # Execute logic for Paragraphs and Tables
                    process_paragraphs(doc.paragraphs)
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                process_paragraphs(cell.paragraphs)
                    
                    out_stream = io.BytesIO()
                    doc.save(out_stream)
                    zip_f.writestr(f.name, out_stream.getvalue())

            st.success(f"Successfully processed {len(uploaded_prison_files)} letters!")
            st.download_button("📥 Download All Updated Letters (ZIP)", zip_buffer.getvalue(), f"Smart_Update_{datetime.now().strftime('%d_%m')}.zip")
