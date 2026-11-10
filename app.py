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
    pdf.add_page()
    
    if pdf.dejavu_loaded:
        pdf.add_font("DejaVu", "", font_path, uni=True)
        pdf.add_font("DejaVu", "B", font_bold_path, uni=True)
        pdf.set_font("DejaVu", "B", 14)
    else:
        pdf.set_font("Helvetica", "B", 14)
        
    pdf.set_text_color(30, 41, 59)
    pdf.cell(0, 10, "Dzivojamas telpas ires liguma juridiska atskaite", ln=True)
    
    if pdf.dejavu_loaded:
        pdf.set_font("DejaVu", "", 9)
    else:
        pdf.set_font("Helvetica", "", 9)
        
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Izveides datums: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}", ln=True)
    pdf.ln(5)
    
    lines = analysis_text.split('\n')
    for line in lines:
        line_str = line.strip()
        if not line_str:
            pdf.ln(3)
            continue
            
        if line_str.startswith("## "):
            pdf.ln(4)
            if pdf.dejavu_loaded:
                pdf.set_font("DejaVu", "B", 11)
            else:
                pdf.set_font("Helvetica", "B", 11)
            
            if "Kritiski" in line_str or "🔴" in line_str:
                pdf.set_text_color(220, 38, 38)
            elif "Brīdinājumi" in line_str or "🟡" in line_str:
                pdf.set_text_color(217, 119, 6)
            elif "Ieteikumi" in line_str or "🟢" in line_str:
                pdf.set_text_color(22, 163, 74)
            else:
                pdf.set_text_color(30, 41, 59)
                
            clean_title = line_str.replace("##", "").strip()
            clean_title_pdf = clean_title.replace("🔴", "[!]").replace("🟡", "[?]").replace("🟢", "[OK]").replace("📊", "[i]")
            pdf.cell(0, 7, clean_title_pdf, ln=True)
            
            if pdf.dejavu_loaded:
                pdf.set_font("DejaVu", "", 9.5)
            else:
                pdf.set_font("Helvetica", "", 9.5)
            pdf.set_text_color(40, 40, 40)
        else:
            clean_line = line_str.replace("🔴", "[!]").replace("🟡", "[?]").replace("🟢", "[OK]").replace("📊", "[i]")
            pdf.multi_cell(0, 5.5, clean_line)
            
    pdf_buffer = io.BytesIO()
    pdf_bytes = pdf.output(dest='S')
    if isinstance(pdf_bytes, str):
        pdf_bytes = pdf_bytes.encode('latin1', errors='replace')
    pdf_buffer.write(pdf_bytes)
    pdf_buffer.seek(0)
    return pdf_buffer

# Poga līguma analīzei
analyze_button = st.button("🔍 Analizēt līgumu", type="primary")

# Analīzes izpilde
if analyze_button:
    if not api_key:
        st.error("⚠️ Lūdzu, ievadi savu Gemini API atslēgu sānu joslā!")
    elif uploaded_file is None:
        st.error("⚠️ Lūdzu, augšupielādē līguma failu!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-3.8-flash')
            
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            with st.spinner("🔄 Notiek līguma analīze un risku izvērtēšana..."):
                if file_extension == 'docx':
                    bytes_data = uploaded_file.getvalue()
                    target_stream = io.BytesIO(bytes_data)
                    doc = docx.Document(target_stream)
                    contract_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                    
                    full_prompt = f"{SYSTEM_PROMPT}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{contract_text}"
                    response = model.generate_content(full_prompt, stream=True)
                    
                else:
                    file_bytes = uploaded_file.getvalue()
                    mime_type_map = {
                        'pdf': 'application/pdf',
                        'png': 'image/png',
                        'jpg': 'image/jpeg',
                        'jpeg': 'image/jpeg'
                    }
                    mime_type = mime_type_map.get(file_extension, 'application/octet-stream')
                    
                    file_part = {
                        'mime_type': mime_type,
                        'data': file_bytes
                    }
                    
                    response = model.generate_content([SYSTEM_PROMPT, file_part], stream=True)
            
            st.markdown("### 📋 Analīzes rezultāts reāllaikā:")
            
            def stream_and_collect():
                for chunk in response:
                    if hasattr(chunk, 'text') and chunk.text:
                        yield chunk.text

            full_response_text = st.write_stream(stream_and_collect())
            
            if full_response_text:
                st.success("✅ Analīze veiksmīgi pabeigta!")
                st.markdown("---")
                
                st.markdown("### 📥 Lejupielādēt glītu atskaiti saviem klientiem:")
                col1, col2 = st.columns(2)
                
                with col1:
                    docx_file = generate_pretty_docx(full_response_text)
                    st.download_button(
                        label="📄 Lejupielādēt Word (.docx)",
                        data=docx_file,
                        file_name=f"Ires_Liguma_Analize_{datetime.date.today()}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                with col2:
                    pdf_file = generate_pretty_pdf(full_response_text)
                    st.download_button(
                        label="📄 Lejupielādēt PDF (.pdf)",
                        data=pdf_file,
                        file_name=f"Ires_Liguma_Analize_{datetime.date.today()}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
        except Exception as e:
            st.error(f"❌ Kļūda apstrādājot failu: {str(e)}")
            
