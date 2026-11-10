import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import io
import datetime
import urllib.request
import os

# Lapas konfigurācija
st.set_page_config(
    page_title="Dzīvokļu un Nomas Līgumu AI Skeneris",
    page_icon="🏠",
    layout="centered"
)

st.title("🏠 Dzīvokļu Īres Līgumu AI Skeneris")
st.markdown("Augšupielādē savu līgumu un saņem tūlītēju reāllaika straumēšanas analīzi.")

# Sānu josla
st.sidebar.header("🔑 Piekļuves iestatījumi")
api_key = st.sidebar.text_input("Ievadi Gemini API atslēgu", type="password")

uploaded_file = st.file_uploader(
    "Izvēlies vai ievelc līguma failu (.docx, .pdf, .png, .jpg):",
    type=["docx", "pdf", "png", "jpg", "jpeg"]
)

SYSTEM_PROMPT = """
Tu esi pieredzējis Latvijas nekustamā īpašuma un īres tiesību eksperts. Tavs uzdevums ir izanalizēt augšupielādēto īres vai nomas līgumu saskaņā ar Latvijas Dzīvojamo telpu īres likumu un Civillikumu. 

Raksti vienkāršā, cilvēcīgā un viegli saprotamā latviešu valodā. Izmanto tieši šādu virsrakstu struktūru ar krāsu indikatoriem:

## 📊 Kopējais novērtējums
[1-2 teikumos pasaki, vai šo līgumu drīkst parakstīt un kāds ir galvenais risks]

## 🔴 Kritiski riski
[Uzskaiti galvenās bīstamās vietas un sodus vienkāršiem vārdiem]

## 🟡 Brīdinājumi
[Norādi uz neskaidriem vai aizdomīgiem punktiem]

## 🟢 Ieteikumi un labojumi
[Skaidri pasaki, kas līgumā ir jāizmaina]
"""

# Funkcija Word dokumenta ģenerēšanai
def generate_pretty_docx(analysis_text):
    doc = docx.Document()
    title = doc.add_heading("ĪRES LĪGUMA JURIDISKĀ ANALĪZE", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_meta = p_meta.add_run(f"Atskaite sagatavota: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}")
    run_meta.font.size = Pt(9)
    run_meta.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    lines = analysis_text.split('\n')
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
            
        if "Kopējais novērtējums" in line_str or "📊" in line_str:
            h = doc.add_heading(line_str.replace("##", "").strip(), level=1)
            h.runs[0].font.color.rgb = RGBColor(30, 41, 59)
        elif "Kritiski riski" in line_str or "🔴" in line_str:
            h = doc.add_heading(line_str.replace("##", "").strip(), level=1)
            h.runs[0].font.color.rgb = RGBColor(220, 38, 38)
        elif "Brīdinājumi" in line_str or "🟡" in line_str:
            h = doc.add_heading(line_str.replace("##", "").strip(), level=1)
            h.runs[0].font.color.rgb = RGBColor(217, 119, 6)
        elif "Ieteikumi" in line_str or "🟢" in line_str:
            h = doc.add_heading(line_str.replace("##", "").strip(), level=1)
            h.runs[0].font.color.rgb = RGBColor(22, 163, 74)
        else:
            p = doc.add_paragraph(line_str)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# PDF ģenerēšanas klase ar drošu faila un teksta apstrādi
class PDFReport(FPDF):
    def header(self):
        self.set_fill_color(30, 41, 59)
        self.rect(0, 0, 210, 15, 'F')
        if hasattr(self, 'dejavu_loaded') and self.dejavu_loaded:
            self.set_font("DejaVu", "B", 10)
        else:
            self.set_font("Helvetica", "B", 10)
        self.set_text_color(255, 255, 255)
        self.set_xy(10, 4)
        self.cell(0, 8, "JURIDISKAIS ATZINUMS — IRES LIGUMA ANALIZE", 0, 0, "L")
        self.ln(20)

    def footer(self):
        self.set_y(-15)
        if hasattr(self, 'dejavu_loaded') and self.dejavu_loaded:
            self.set_font("DejaVu", "", 8)
        else:
            self.set_font("Helvetica", "", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Lappuse {self.page_no()} | AI Juridiskais Asistents", 0, 0, "C")

def generate_pretty_pdf(analysis_text):
    font_path = "DejaVuSans.ttf"
    font_bold_path = "DejaVuSans-Bold.ttf"
    
    if not os.path.exists(font_path):
        try:
            urllib.request.urlretrieve("https://github.com/dejavu-fonts/dejavu-fonts-ttf/raw/master/ttf/DejaVuSans.ttf", font_path)
        except Exception:
            pass
            
    if not os.path.exists(font_bold_path):
        try:
            urllib.request.urlretrieve("https://github.com/dejavu-fonts/dejavu-fonts-ttf/raw/master/ttf/DejaVuSans-Bold.ttf", font_bold_path)
        except Exception:
            pass

    pdf = PDFReport()
    pdf.dejavu_loaded = os.path.exists(font_path) and os.path.exists(font_bold_path)
    pdf
