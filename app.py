import streamlit as st
from docxtpl import DocxTemplate
import io
import zipfile
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
tab1, tab2 = st.tabs(["📝 Individual Generator", "📂 Bulk Prison Updater (Thai Support)"])

# ==========================================
# TAB 1: INDIVIDUAL DOCUMENT GENERATOR (English)
# ==========================================
with tab1:
    context = {}
    template_file = ""

    st.subheader("🟢 STEP 1: CHOOSE MAIN CATEGORY")
    category = st.radio(
        "Select the department for this document:",
        ["Visa", "Land Transport", "Visa Transfer"],
        horizontal=True, key="indiv_cat_radio"
    )

    st.write("---")
    st.subheader(f"📝 STEP 2: ENTER {category.upper()} DETAILS")

    doc_date = st.date_input("Document Date", value=datetime.now(), key="indiv_date")
    formatted_date = doc_date.strftime("%d %B %Y")

    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Full Name (As shown in Passport)", key="name_in")
        passport = st.text_input("Passport Number", key="pass_in")
    with col2:
        dob = st.text_input("Date of Birth (e.g., 01 Jan 1990)", key="dob_in")
        pob = st.text_input("Place of Birth (e.g., Lagos, Nigeria)", key="pob_in")

    gender_choice = st.radio("Gender of Applicant", ["Male", "Female"], horizontal=True, key="gender_in")
    g1, g2, g3 = ("he", "his", "him") if gender_choice == "Male" else ("she", "her", "her")

    if category == "Visa":
        sub_visa = st.selectbox("Purpose", ["30 Days Extension", "Student", "Employment", "Marriage"])
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

    if st.button("💾 GENERATE DOCUMENT", key="btn_gen"):
        if not name.strip() or not passport.strip():
            st.error("Missing Full Name or Passport.")
        else:
            try:
                doc_t = DocxTemplate(template_file)
                doc_t.render(final_context)
                bio = io.BytesIO()
                doc_t.save(bio)
                st.balloons()
                st.download_button("📥 Download", bio.getvalue(), f"{name}.docx")
            except Exception as e:
                st.error(f"Error: {e}")

# ==========================================
# TAB 2: BULK PRISON UPDATER (Thai Support & Angsana)
# ==========================================
with tab2:
    st.subheader("⚡ Thai Letter Bulk Date Swapper")
    st.info("Upload Thai documents. Replaces dates while forcing font back to **Angsana New (16pt)**.")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**🔍 FIND (OLD)**")
        old_i = st.text_input("Old Issue Date (Thai or English)", placeholder="เช่น 1 พฤษภาคม 2569", key="ti_old_i")
        old_v = st.text_input("Old Visit Date/Month", placeholder="เช่น พฤษภาคม 2569", key="ti_old_v")
    
    with col_b:
        st.markdown("**🖋️ REPLACE (NEW)**")
        new_i = st.text_input("New Issue Date", placeholder="เช่น 1 มิถุนายน 2569", key="ti_new_i")
        new_v = st.text_input("New Visit Date/Month", placeholder="เช่น มิถุนายน 2569", key="ti_new_v")

    files = st.file_uploader("Upload Thai .docx files", type=["docx"], accept_multiple_files=True, key="bulk_up")

    if st.button("🚀 BATCH UPDATE ALL THAI LETTERS", key="btn_bulk"):
        if not files or not old_i or not new_i:
            st.warning("Please enter the 'Find' and 'Replace' dates and upload files.")
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_f:
                for f in files:
                    doc = Document(f)
                    
                    def thai_safe_replace(paragraphs, old_txt, new_txt):
                        for p in paragraphs:
                            if old_txt in p.text:
                                # Replace the text
                                updated_text = p.text.replace(old_txt, new_txt)
                                # Clear and rewrite to fix run splits
                                p.text = ""
                                run = p.add_run(updated_text)
                                # Apply Angsana New for both Latin and Complex (Thai) script
                                run.font.name = 'Angsana New'
                                run.font.size = Pt(16)
                                r = run._element.rPr
                                r.get_or_add_rFonts().set(qn('w:eastAsia'), 'Angsana New')
                                r.get_or_add_rFonts().set(qn('w:cs'), 'Angsana New')
                                r.get_or_add_rFonts().set(qn('w:ascii'), 'Angsana New')
                                r.get_or_add_rFonts().set(qn('w:hAnsi'), 'Angsana New')

                    # Execute replacements
                    thai_safe_replace(doc.paragraphs, old_i, new_i)
                    if old_v: thai_safe_replace(doc.paragraphs, old_v, new_v)
                    
                    # Also check inside tables
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                thai_safe_replace(cell.paragraphs, old_i, new_i)
                                if old_v: thai_safe_replace(cell.paragraphs, old_v, new_v)
                    
                    out = io.BytesIO()
                    doc.save(out)
                    zip_f.writestr(f.name, out.getvalue())

            st.success("Batch update complete!")
            st.download_button("📥 Download Updated Thai ZIP", zip_buffer.getvalue(), f"Updated_Thai_Letters.zip")
