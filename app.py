import streamlit as st
from docxtpl import DocxTemplate
import io
import zipfile
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

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

# --- SMART FONT ENGINE (Enforces Font Rules, Superscript, and Tab Alignment) ---
def apply_smart_font(paragraph, text, is_bold=False, is_underline=False, is_header_block=False):
    # If it's the Ref/Date block, we use a Tab Stop to align them together on the right
    if is_header_block:
        paragraph.paragraph_format.left_indent = None
        paragraph.paragraph_format.first_line_indent = None
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT # Left align the text itself
        # Add a tab stop at 4.5 inches (Adjustable) to push the start point to the right
        tab_stops = paragraph.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Inches(4.5), WD_TAB_ALIGNMENT.LEFT)
        text = "\t" + text # Add the tab character to jump to the start point

    # Split text to catch suffixes for superscripting (st, nd, rd, th)
    parts = re.split(r'(\d+)(st|nd|rd|th)', text)
    
    for part in parts:
        if not part: continue
        run = paragraph.add_run(part
