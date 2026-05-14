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

# --- REFINED FONT & SCRIPT HELPER ---
def apply_smart_font(paragraph, text, is_bold=False, is_underline=False):
    """
    Splits text into chunks.
    Thai & Numbers -> Angsana New 17pt
    English Letters (A-Z) -> Times New Roman 13.5pt
    """
    # Regex splits by English letters vs everything else (Thai, Numbers, Punctuation)
    chunks = re.findall(r'[a-zA-Z]+|[^a-zA-Z]+', text)
    
    for chunk in chunks:
        run = paragraph.add_run(chunk)
        run.bold = is_bold
        run.underline = is_underline
        
        # Check if chunk contains English letters
        if re.search(r'[a-zA-Z]', chunk):
            # English Letters -> Times New Roman 13.5
            run.font.name = 'Times New Roman'
            run.font.size = Pt(13.5)
            r = run._element.rPr
            r.get_or_add_rFonts().set(qn('w:ascii'), 'Times New Roman')
            r.get_or_add_rFonts().set(qn('w:hAnsi'), 'Times New Roman')
        else:
            # Thai, Numbers, Symbols, and Spaces -> Angsana New 17
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
    st.subheader("🟢 STEP 1: CHOOSE CATEGORY")
    category = st.radio("Department:", ["Visa", "Land Transport", "Visa Transfer"], horizontal=True, key="t1_cat")

    st.write("---")
    doc_date = st.date_input("Document Date", value=datetime.now(), key="t1_date")
    formatted_date = doc_date.strftime("%d %B %Y")

    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Full Name", key="t1_name")
        passport = st.text_input("Passport Number", key="t1_pp")
    with col2:
        dob = st.text_input("Date of Birth", key="t1_dob")
        pob = st.text_input("Place of Birth", key="t1_pob")

    gender_choice = st.radio("Gender", ["Male", "Female"], horizontal=True, key="t1_gender")
    g1, g2, g3 = ("he", "his", "him") if gender_choice == "Male" else ("she", "her", "her")

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

    if st.button("💾 GENERATE SINGLE LETTER", key="t1_btn"):
        try:
            doc = DocxTemplate(template_file)
            doc.render(final_context)
            bio = io.BytesIO()
            doc.save(bio)
            st.success(f"Generated for {name}")
            st.download_button("📥 Download", bio.getvalue(), f"{name}.docx")
        except Exception as e:
            st.error(f"Error: {e}. Ensure {template_file} is in GitHub.")

# ==========================================
# TAB 2: PROFESSIONAL BULK UPDATER
# ==========================================
with tab2:
    st.subheader("🏛️ Smart Prison Batch Updater")
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**📅 DATE OF ISSUE SETUP**")
        ref_id = st.text_input("Search for Reference line:", value="อีเอ็นบี/ซีเอ็น.07")
        new_issue = st.text_input("New Date of Issue", value=datetime.now().strftime("%d %B %Y"))
    
    with col_b:
        st.markdown("**🖋️ VISIT DETAILS SETUP**")
        new_visit = st.text_input("New Month & Time Range", placeholder="มิถุนายน 2569 เวลา 12.00 - 15.00 น.")

    files = st.file_uploader("Upload Word Docs", type=["docx"], accept_multiple_files=True)

    if st.button("🚀 BATCH UPDATE ALL DOCUMENTS"):
        if not files or not new_visit:
            st.error("Please upload files and provide visit details.")
        else:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zip_f:
                for f in files:
                    doc = Document(f)
                    update_next = False
                    
                    for p in list(doc.paragraphs):
                        # 1. Update Date Below Reference No.
                        if ref_id in p.text:
                            update_next = True
                            continue
                        if update_next and len(p.text.strip()) > 0:
                            p.text = ""
                            apply_smart_font(p, new_issue)
                            update_next = False
                            continue
                        
                        # 2. Update Details after "ในเดือน"
                        if "ในเดือน" in p.text:
                            parts = p.text.split("ในเดือน", 1)
                            p.text = ""
                            # "ในเดือน" and preceding text (Normal)
                            apply_smart_font(p, parts[0] + "ในเดือน")
                            # Space after "ในเดือน" (Normal)
                            apply_smart_font(p, " ")
                            # New Visit Info (Bold + Underlined)
                            apply_smart_font(p, new_visit, is_bold=True, is_underline=True)

                    # 3. Check Tables
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
            
            st.success("Batch Complete! Numbers & Thai are Angsana 17. English is Times 13.5.")
            st.download_button("📥 Download Updated ZIP", zip_buf.getvalue(), "Embassy_Prison_Updated.zip")
