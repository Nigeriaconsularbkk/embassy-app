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

# --- FONT & STYLE HELPER FUNCTION ---
def apply_mixed_style(run, text, is_bold=False, is_underline=False):
    """Applies Angsana New 17 for Thai and Times New Roman 13.5 for English."""
    run.text = text
    run.bold = is_bold
    run.underline = is_underline
    
    # Base size 17 (Angsana looks smaller than Times, so we keep 17 as primary)
    run.font.size = Pt(17)
    
    # Set Thai/Complex Script
    run.font.name = 'Angsana New'
    r = run._element.rPr
    r.get_or_add_rFonts().set(qn('w:cs'), 'Angsana New')
    r.get_or_add_rFonts().set(qn('w:eastAsia'), 'Angsana New')
    
    # Set English/Ascii Script
    r.get_or_add_rFonts().set(qn('w:ascii'), 'Times New Roman')
    r.get_or_add_rFonts().set(qn('w:hAnsi'), 'Times New Roman')

# --- OFFICIAL HEADER ---
st.markdown("<h1 class='main-title'>EMBASSY OF NIGERIA BKK SYSTEM</h1>", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📝 Individual Generator", "📂 Professional Bulk Updater"])

# ==========================================
# TAB 1: INDIVIDUAL DOCUMENT GENERATOR
# ==========================================
with tab1:
    st.subheader("🟢 STEP 1: CHOOSE CATEGORY")
    category = st.radio("Department:", ["Visa", "Land Transport", "Visa Transfer"], horizontal=True)

    st.write("---")
    doc_date = st.date_input("Document Date", value=datetime.now())
    formatted_date = doc_date.strftime("%d %B %Y")

    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Full Name")
        passport = st.text_input("Passport Number")
    with col2:
        dob = st.text_input("Date of Birth")
        pob = st.text_input("Place of Birth")

    gender_choice = st.radio("Gender", ["Male", "Female"], horizontal=True)
    g1, g2, g3 = ("he", "his", "him") if gender_choice == "Male" else ("she", "her", "her")

    # Category Logic
    context = {}
    template_file = ""
    if category == "Visa":
        sub = st.selectbox("Purpose", ["30 Days Extension", "Student", "Employment", "Marriage"])
        template_file = f"visa_{sub.lower().replace(' ', '')}.docx"
    elif category == "Land Transport":
        template_file = "land_transport.docx"
    elif category == "Visa Transfer":
        template_file = "visa_transfer.docx"

    final_context = {
        "name": name, "name_capital": name.upper() if name else "",
        "passport": passport, "dob": dob, "pob": pob,
        "gender1": g1, "gender2": g2, "gender3": g3, "date": formatted_date
    }

    if st.button("💾 GENERATE SINGLE LETTER"):
        try:
            doc = DocxTemplate(template_file)
            doc.render(final_context)
            bio = io.BytesIO()
            doc.save(bio)
            st.success(f"Generated for {name}")
            st.download_button("📥 Download", bio.getvalue(), f"{name}.docx")
        except Exception as e:
            st.error(f"Error: {e}. Make sure {template_file} is in your GitHub folder.")

# ==========================================
# TAB 2: PROFESSIONAL BULK UPDATER
# ==========================================
with tab2:
    st.subheader("🏛️ Smart Prison Batch Updater")
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**📅 DATE OF ISSUE**")
        ref_id = st.text_input("Search for Ref No:", value="อีเอ็นบี/ซีเอ็น.07")
        new_issue = st.text_input("New Date of Issue", value=datetime.now().strftime("%d %B %Y"))
    
    with col_b:
        st.markdown("**🖋️ VISIT DETAILS**")
        new_visit = st.text_input("New Month & Time Range", placeholder="มิถุนายน 2569 เวลา 12.00 - 15.00 น.")

    files = st.file_uploader("Upload Word Docs", type=["docx"], accept_multiple_files=True)

    if st.button("🚀 BATCH UPDATE ALL"):
        if not files or not new_visit:
            st.error("Upload files and fill visit details.")
        else:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zip_f:
                for f in files:
                    doc = Document(f)
                    update_next = False
                    
                    for p in list(doc.paragraphs):
                        # Issue Date Logic
                        if ref_id in p.text:
                            update_next = True
                            continue
                        if update_next and len(p.text.strip()) > 0:
                            p.text = ""
                            apply_mixed_style(p.add_run(new_issue), new_issue)
                            update_next = False
                            continue
                        
                        # "ในเดือน" Logic (Bold + Underline)
                        if "ในเดือน" in p.text:
                            parts = p.text.split("ในเดือน", 1)
                            p.text = ""
                            apply_mixed_style(p.add_run(parts[0] + "ในเดือน"), parts[0] + "ในเดือน")
                            apply_mixed_style(p.add_run(" " + new_visit), " " + new_visit, is_bold=True, is_underline=True)

                    # Check Tables
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    if "ในเดือน" in p.text:
                                        parts = p.text.split("ในเดือน", 1)
                                        p.text = ""
                                        apply_mixed_style(p.add_run(parts[0] + "ในเดือน"), parts[0] + "ในเดือน")
                                        apply_mixed_style(p.add_run(" " + new_visit), " " + new_visit, is_bold=True, is_underline=True)

                    out = io.BytesIO()
                    doc.save(out)
                    zip_f.writestr(f.name, out.getvalue())
            
            st.success("Batch Complete!")
            st.download_button("📥 Download ZIP", zip_buf.getvalue(), "Embassy_Batch.zip")
