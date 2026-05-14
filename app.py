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

# --- DATE ORDINAL HELPER ---
def get_ordinal_suffix(day):
    if 11 <= day <= 13:
        return 'th'
    else:
        return {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')

# --- CUSTOM GREEN THEME CSS ---
st.markdown("""
    <style>
    .main-title { color: #008751; text-align: center; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #008751; padding-bottom: 10px; }
    .stSubheader { color: #008751 !important; font-weight: bold; }
    div.stButton > button:first-child { background-color: #008751; color: white; border-radius: 10px; border: none; height: 3em; width: 100%; font-weight: bold; font-size: 20px; }
    [data-testid="stSidebar"] { background-color: #f0f2f6; border-right: 2px solid #008751; }
    </style>
    """, unsafe_allow_html=True)

# --- SMART FONT ENGINE (Ensures Alignment and Font Rules) ---
def apply_smart_font(paragraph, text, is_bold=False, is_underline=False):
    # Ensure the paragraph has no weird offsets so lines align vertically
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.first_line_indent = None
    
    # Split text to catch suffixes for superscripting (st, nd, rd, th)
    parts = re.split(r'(\d+)(st|nd|rd|th)', text)
    
    for part in parts:
        if not part: continue
        run = paragraph.add_run(part)
        run.bold = is_bold
        run.underline = is_underline
        
        # 1. Handle Superscript for dates
        if part in ['st', 'nd', 'rd', 'th']:
            run.font.superscript = True
            run.font.size = Pt(10)
        
        # 2. Font Rules: Thai/Numbers = Angsana New, English = Times New Roman
        if re.search(r'[\u0e00-\u0e7f0-9]', part):
            run.font.name = 'Angsana New'
            run.font.size = Pt(17)
            r = run._element.rPr
            r.get_or_add_rFonts().set(qn('w:cs'), 'Angsana New')
            r.get_or_add_rFonts().set(qn('w:eastAsia'), 'Angsana New')
            r.get_or_add_rFonts().set(qn('w:ascii'), 'Angsana New')
            r.get_or_add_rFonts().set(qn('w:hAnsi'), 'Angsana New')
        else:
            run.font.name = 'Times New Roman'
            if not run.font.superscript:
                run.font.size = Pt(13.5)
            r = run._element.rPr
            r.get_or_add_rFonts().set(qn('w:ascii'), 'Times New Roman')
            r.get_or_add_rFonts().set(qn('w:hAnsi'), 'Times New Roman')

# --- TOP PRIORITY: DATE SELECTION (SIDEBAR) ---
st.sidebar.header("📅 GLOBAL SETTINGS")
today_picker = st.sidebar.date_input("Today's Date", value=datetime.now())
suffix = get_ordinal_suffix(today_picker.day)
# Format: 14th May, 2026
formatted_date_plain = f"{today_picker.day}{suffix} {today_picker.strftime('%B, %Y')}"

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
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Full Name")
        passport = st.text_input("Passport Number")
        dob = st.text_input("Date of Birth")
    with col2:
        pob = st.text_input("Place of Birth")
        gender_choice = st.radio("Gender", ["Male", "Female"], horizontal=True)

    g1, g2, g3 = ("he", "his", "him") if gender_choice == "Male" else ("she", "her", "her")

    final_context = {
        "name": name, "name_capital": name.upper() if name else "",
        "passport": passport, "dob": dob, "pob": pob,
        "gender1": g1, "gender2": g2, "gender3": g3,
        "gender": g1, "date": formatted_date_plain
    }

    template_file = ""
    if category == "Visa 30 Days Extension":
        template_file = "visa_30days.docx"
        final_context["leave_on"] = st.text_input("Intended Leave Date")
    elif category == "Visa Student":
        template_file = "visa_student.docx"
        final_context.update({"program": st.text_input("Program"), "place_of_study": st.text_input("University"), "location_of_study": st.text_input("Location")})
    elif category == "Visa Employment":
        template_file = "visa_employment.docx"
        final_context.update({"place_of_work": st.text_input("Company"), "location_of_work": st.text_input("Location"), "place_of_issue": st.text_input("Passport Place of Issue"), "country_of_issue": "Nigeria", "date_of_issue": st.text_input("Passport Issue Date"), "passport_expiration": st.text_input("Passport Expiry Date")})
    elif category == "Visa Marriage":
        template_file = "visa_marriage.docx"
    elif category == "Land Transport":
        template_file = "land_transport.docx"
        final_context["current_address"] = st.text_area("Address")
        final_context["purpose"] = st.selectbox("Purpose:", ["registering a driving license as requested", "transferring a vehicle as requested"])
    elif category == "Visa Transfer":
        template_file = "visa_transfer.docx"
        final_context.update({"old_passport": st.text_input("Old Passport No."), "old_passport_expiration": st.text_input("Old Passport Expiry"), "place_of_issue": st.text_input("New Passport Place of Issue"), "date_of_issue": st.text_input("New Passport Issue Date"), "passport_expiration": st.text_input("New Passport Expiry")})

    if st.button("💾 GENERATE DOCUMENT"):
        if name and passport:
            try:
                doc = DocxTemplate(template_file)
                doc.render(final_context)
                target_stream = io.BytesIO()
                doc.save(target_stream)
                target_stream.seek(0)
                final_doc = Document(target_stream)
                for p in final_doc.paragraphs:
                    if formatted_date_plain in p.text:
                        orig_text = p.text
                        p.text = "" 
                        apply_smart_font(p, orig_text)
                final_bio = io.BytesIO()
                final_doc.save(final_bio)
                st.success(f"Generated for {formatted_date_plain}")
                st.download_button("📥 Download", final_bio.getvalue(), f"{name}_{category}.docx")
            except Exception as e:
                st.error(f"Error: {e}")

# ==========================================
# TAB 2: PROFESSIONAL BULK UPDATER
# ==========================================
with tab2:
    st.subheader("🏛️ Smart Prison Batch Updater")
    col_a, col_b = st.columns(2)
    with col_a:
        ref_id = st.text_input("Reference No. to Search:", value="อีเอ็นบี/ซีเอ็น.07")
        new_issue = st.text_input("New Issue Date:", value=formatted_date_plain)
    with col_b:
        new_visit = st.text_input("New Visit Details:", placeholder="มิถุนายน 2569")

    files = st.file_uploader("Upload Word Docs", type=["docx"], accept_multiple_files=True)

    if st.button("🚀 BATCH UPDATE"):
        if files and new_visit:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zip_f:
                for f in files:
                    doc = Document(f)
                    update_next_issue = False
                    
                    for p in list(doc.paragraphs):
                        # 1. Find Reference line, then update the very next paragraph (the date line)
                        if ref_id in p.text:
                            # Align the reference line itself just in case
                            p.paragraph_format.left_indent = None
                            p.paragraph_format.first_line_indent = None
                            update_next_issue = True
                            continue
                            
                        if update_next_issue and len(p.text.strip()) > 0:
                            p.text = ""
                            # This will write the new date with 0 indentation, perfectly aligned with the ref above
                            apply_smart_font(p, new_issue)
                            update_next_issue = False
                            continue
                        
                        # 2. Update visit details paragraph
                        if "ในเดือน" in p.text:
                            parts = p.text.split("ในเดือน", 1)
                            p.text = ""
                            apply_smart_font(p, parts[0] + "ในเดือน")
                            apply_smart_font(p, " ")
                            apply_smart_font(p, new_visit, is_bold=True, is_underline=True)
                        elif p.text.strip():
                            # Standardize font for everything else in the document
                            orig_p_text = p.text
                            p.text = ""
                            apply_smart_font(p, orig_p_text)
                    
                    out = io.BytesIO()
                    doc.save(out)
                    zip_f.writestr(f.name, out.getvalue())
            st.download_button("📥 Download ZIP", zip_buf.getvalue(), "Updated_Prison_Files.zip")
