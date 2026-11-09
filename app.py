import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt, RGBColor
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
st.markdown("Augšupielādē savu Word (.docx) līgumu un saņem tūlītēju analīzi un augstākās kvalitātes atskaites.")

# Sānu josla
st.sidebar.header("🔑 Piekļuves iestatījumi")
api_key = st.sidebar.text_input("Ievadi Gemini API atslēgu", type="password")

uploaded_file = st.file_uploader("Izvēlies vai ievelc savu līguma Word failu (.docx):", type=["docx"])

SYSTEM_PROMPT = """
Tu esi pieredzējis Latvijas nekustamā īpašuma un īres tiesību eksperts. Tavs uzdevums ir profesionāli izanalizēt īres līgumu saskaņā ar Latvijas Dzīvojamo telpu īres likumu un Civillikumu.

Formatē savu atbildi precīzi ar šādām sadaļām:

## 📊 Kopējais novērtējums
[Īss kopsavilkums 1-2 teikumos par līguma drošību un kopējo riska līmeni]

## 🔴 Kritiski riski
[Uzskaitījums ar bīstamajiem punktiem, nesamērīgiem sodiem vai likumpārkāpumiem]

## 🟡 Brīdinājumi
[Uzskaitījums ar neskaidriem vai īrniekam nelabvēlīgiem nosacījumiem]

## 🟢 Ieteikumi un labojumi
[Konkrēti teksta labojumi un juridiski korekti formulējumi]
"""

# Funkcija glīta Word dokumenta ģenerēšanai
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
            
        if line_str.startswith("## 📊"):
            h = doc.add_heading("📊 Kopējais novērtējums", level=1)
            h.runs[0].font.color.rgb = RGBColor(30, 41, 59)
        elif line_str.startswith("## 🔴"):
            h = doc.add_heading("🔴 Kritiski riski un bīstamie panti", level=1)
            h.runs[0].font.color.rgb = RGBColor(220, 38, 38)
        elif line_str.startswith("## 🟡"):
            h = doc.add_heading("🟡 Brīdinājumi un slēptie āķi", level=1)
            h.runs[0].font.color.rgb = RGBColor(217, 119, 6)
        elif line_str.startswith("## 🟢"):
            h = doc.add_heading("🟢 Zaļie ieteikumi un labojumi", level=1)
            h.runs[0].font.color.rgb = RGBColor(22, 163, 74)
        else:
            p = doc.add_paragraph(line_str)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# Profesionāla PDF ģenerēšanas klase ar Unicode un latviešu valodas atbalstu
class PDFReport(FPDF):
    def header(self):
        # Augšējā dekoratīvā līnija un virsraksts
        self.set_fill_color(30, 41, 59) # Tumši zils/pelēks tonis
        self.rect(0, 0, 210, 15, 'F')
        self.set_font("DejaVu", "B", 10)
        self.set_text_color(255, 255, 255)
        self.set_xy(10, 4)
        self.cell(0, 8, "JURIDISKAIS ATZINUMS — ĪRES LĪGUMA ANALĪZE", 0, 0, "L")
        self.ln(20)

    def footer(self):
        # Apakšējā līnija un lapas numurs
        self.set_y(-15)
        self.set_font("DejaVu", "", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Lappuse {self.page_no()} | Sagatavots ar AI Juridisko Asistentu", 0, 0, "C")

def generate_pretty_pdf(analysis_text):
    # Lejupielādējam Unicode fontu latviešu burtiem, ja tāda vēl nav
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
    pdf.add_page()
    
    # Pievienojam Unicode fontu
    if os.path.exists(font_path):
        pdf.add_font("DejaVu", "", font_path, uni=True)
    if os.path.exists(font_bold_path):
        pdf.add_font("DejaVu", "B", font_bold_path, uni=True)
    
    # Dokumenta pamatinformācija
    pdf.set_font("DejaVu", "B", 14)
    pdf.set_text_color(30, 41, 59)
    pdf.cell(0, 10, "Dzīvojamās telpas īres līguma juridiskā risku atskaite", ln=True)
    
    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Izveides datums: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}", ln=True)
    pdf.ln(5)
    
    # Apstrādājam analīzes tekstu pa rindām
    lines = analysis_text.split('\n')
    for line in lines:
        line_str = line.strip()
        if not line_str:
            pdf.ln(3)
            continue
            
        # Pārbaudām virsrakstus un piešķiram tiem atbilstošas krāsas
        if line_str.startswith("## "):
            pdf.ln(4)
            pdf.set_font("DejaVu", "B", 11)
            
            if "Kritiski riski" in line_str or "🔴" in line_str:
                pdf.set_text_color(220, 38, 38) # Sarkans
            elif "Brīdinājumi" in line_str or "🟡" in line_str:
                pdf.set_text_color(217, 119, 6) # Oranžs
            elif "Ieteikumi" in line_str or "🟢" in line_str:
                pdf.set_text_color(22, 163, 74) # Zaļš
            else:
                pdf.set_text_color(30, 41, 59)  # Pārējie
                
            clean_title = line_str.replace("## ", "").replace("📊", "").replace("🔴", "").replace("🟡", "").replace("🟢", "").strip()
            pdf.cell(0, 7, clean_title, ln=True)
            pdf.set_font("DejaVu", "", 9.5)
            pdf.set_text_color(40, 40, 40)
        else:
            # Parasts rindkopas teksts ar pilnu latviešu valodas atbalstu
            pdf.multi_cell(0, 5.5, line_str)
            
    pdf_buffer = io.BytesIO()
    pdf_output = pdf.output(dest='S')
    pdf_buffer.write(pdf_output)
    pdf_buffer.seek(0)
    return pdf_buffer

# Analīzes izpilde
if st.button("🔍 Analizēt līgumu tagad", type="primary"):
    if not api_key:
        st.error("⚠️ Lūdzu, ievadi sev Gemini API atslēgu sānu joslā!")
    elif uploaded_file is None:
        st.error("⚠️ Lūdzu, augšupielādē Word (.docx) failu analīzei!")
    else:
        try:
            bytes_data = uploaded_file.getvalue()
            target_stream = io.BytesIO(bytes_data)
            doc = docx.Document(target_stream)
            contract_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            
            if not contract_text.strip():
                st.warning("⚠️ Fails ir tukšs.")
            else:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-1.5-flash')
                full_prompt = f"{SYSTEM_PROMPT}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{contract_text}"
                
                st.markdown("### 📋 Analīzes rezultāts:")
                
                # Reāllaika atbildes izvadīšana
                response = model.generate_content(full_prompt, stream=True)
                full_response_text = st.write_stream(response)
                
                st.success("✅ Analīze veiksmīgi pabeigta!")
                st.markdown("---")
                
                # Lejupielādes pogas
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
