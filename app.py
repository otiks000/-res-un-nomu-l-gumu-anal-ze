import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import io
import re
import datetime
import os

# ---------------------------------------------------------------
# Lapas konfigurācija
# ---------------------------------------------------------------
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

Raksti vienkāršā, cilvēcīgā un viegli saprotamā latviešu valodā.

DARBA METODE:
- Pārskati līgumu punktu pa punktam, no sākuma līdz beigām, un neizlaid nevienu būtisku punktu.
- Esi rūpīgs un kritisks. Labāk atzīmē vairāk potenciālu problēmu nekā mazāk.
- Katram punktam norādi līguma punkta numuru un vienkāršu paskaidrojumu, kāpēc tas ir problēma.

KLASIFIKĀCIJA (katrs punkts drīkst būt tikai VIENĀ sadaļā):

🔴 KRITISKS RISKS - punkti, kas var radīt nopietnus finansiālus vai juridiskus zaudējumus:
- noteikumi, kas ir pretrunā ar likumu vai var būt spēkā neesoši;
- vienpusēja līguma izbeigšana, vienpusēja īres maksas maiņa vai izlikšana bez pamatojuma;
- depozīta paturēšana vai atmaksas nosacījumi, kas ļauj izīrētājam to patvaļīgi neatmaksāt;
- pārmērīgi līgumsodi, soda naudas vai atbildība bez vainas;
- svarīgu aizsardzības mehānismu trūkums (piemēram, nav pieņemšanas-nodošanas akta, nav noteikta kārtība strīdu gadījumā, nav tiesību aizsardzības pārdošanas gadījumā);
- jebkas cits, kas var novest pie lielas naudas zaudēšanas vai mājokļa zaudēšanas.

🟡 BRĪDINĀJUMS - punkti, kas nav tieši nelikumīgi, bet var radīt strīdu vai neskaidrību:
- neskaidri formulējumi (piemēram, "sīkie remontdarbi", "citi maksājumi");
- trūkstošas summas, termiņi vai kārtība;
- nosacījumi, kas ir vienpusēji izdevīgi vienai pusei;
- komunālo maksājumu, uzturēšanas vai remonta sadalījums, kas nav skaidri definēts.

🟢 IETEIKUMS - konkrēts labojums vai papildinājums ar piedāvātu jaunu formulējumu.

Pirms atbildes pārbaudi, lai neviens punkts neatkārtotos vairākās sadaļās.
Ja kādā sadaļā tiešām nav problēmu, to godīgi pasaki, bet vispirms vēlreiz pārbaudi visu līgumu.

Izmanto tieši šādu virsrakstu struktūru:

## 📊 Kopējais novērtējums
[1-2 teikumos pasaki, vai šo līgumu drīkst parakstīt un kāds ir galvenais risks]

## 🔴 Kritiski riski
[Punkti ar līguma punkta numuru un paskaidrojumu]

## 🟡 Brīdinājumi
[Punkti ar līguma punkta numuru un paskaidrojumu]

## 🟢 Ieteikumi un labojumi
[OBLIGĀTI: katram punktam, kas minēts sadaļās "Kritiski riski" un "Brīdinājumi", jābūt atbilstošam ieteikumam ar to pašu līguma punkta numuru un konkrētu piedāvāto jauno formulējumu pēdiņās. Neviens atzīmētais punkts nedrīkst palikt bez ieteikuma.]

PĒDĒJĀ PĀRBAUDE PIRMS ATBILDES: saskaiti punktus sadaļās "Kritiski riski" un "Brīdinājumi" un pārliecinies, ka sadaļā "Ieteikumi un labojumi" ir ieteikums katram no tiem. Ja kāds trūkst, pievieno to.
"""


# ---------------------------------------------------------------
# Palīgfunkcijas teksta tīrīšanai
# ---------------------------------------------------------------
def strip_emoji(text):
    """Noņem emocijzīmes (DejaVu fonts tās neatbalsta PDF failā)."""
    text = re.sub(r"[\U00010000-\U0010FFFF]", "", text)
    text = re.sub(r"[\u2600-\u27BF\uFE0F]", "", text)
    return text


def clean_markdown(text):
    """Noņem markdown zīmes (** un __), lai tās neparādās dokumentos."""
    return text.replace("**", "").replace("__", "")


# ---------------------------------------------------------------
# WORD dokumenta ģenerēšana
# ---------------------------------------------------------------
def generate_pretty_docx(analysis_text):
    doc = docx.Document()
    title = doc.add_heading("ĪRES LĪGUMA JURIDISKĀ ANALĪZE", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_meta = p_meta.add_run(
        f"Atskaite sagatavota: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}"
    )
    run_meta.font.size = Pt(9)
    run_meta.font.color.rgb = RGBColor(100, 100, 100)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    for line in analysis_text.split("\n"):
        line_str = line.strip()
        if not line_str:
            continue

        if line_str.startswith("#"):
            heading_text = strip_emoji(clean_markdown(line_str.lstrip("#"))).strip()
            h = doc.add_heading(heading_text, level=1)

            if "Kritiski" in line_str or "🔴" in line_str:
                color = RGBColor(220, 38, 38)
            elif "Brīdinājumi" in line_str or "🟡" in line_str:
                color = RGBColor(217, 119, 6)
            elif "Ieteikumi" in line_str or "🟢" in line_str:
                color = RGBColor(22, 163, 74)
            else:
                color = RGBColor(30, 41, 59)

            if h.runs:
                h.runs[0].font.color.rgb = color
        else:
            p = doc.add_paragraph(clean_markdown(line_str))
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# ---------------------------------------------------------------
# PDF ģenerēšana (ar pilnu latviešu garumzīmju atbalstu)
# ---------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Vietas, kur meklēt DejaVu fontus (pirmā atrastā tiks izmantota)
FONT_CANDIDATES = [
    # Sistēmas fonti (instalējas ar packages.txt -> fonts-dejavu-core)
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    # Fonti, kas ielikti pašā repozitorijā mapē "fonts"
    (os.path.join(BASE_DIR, "fonts", "DejaVuSans.ttf"),
     os.path.join(BASE_DIR, "fonts", "DejaVuSans-Bold.ttf")),
]


def find_fonts():
    for regular, bold in FONT_CANDIDATES:
        if os.path.exists(regular) and os.path.exists(bold):
            return regular, bold
    return None, None


class PDFReport(FPDF):
    def header(self):
        self.set_fill_color(30, 41, 59)
        self.rect(0, 0, 210, 15, "F")
        self.set_font("DejaVu", "B", 10)
        self.set_text_color(255, 255, 255)
        self.set_xy(10, 4)
        self.cell(0, 8, "JURIDISKAIS ATZINUMS — ĪRES LĪGUMA ANALĪZE", align="L")
        self.set_y(22)

    def footer(self):
        self.set_y(-15)
        self.set_font("DejaVu", "", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Lappuse {self.page_no()} | AI Juridiskais Asistents", align="C")


def generate_pretty_pdf(analysis_text):
    font_path, font_bold_path = find_fonts()
    if not font_path:
        raise FileNotFoundError(
            "Nav atrasti DejaVu fonti. Pārbaudi, vai repozitorijā ir fails packages.txt "
            "ar rindu 'fonts-dejavu-core' un pārstartē lietotni (Reboot app)."
        )

    pdf = PDFReport()
    # Fontus jāreģistrē PIRMS add_page(), jo header() tos izmanto
    pdf.add_font("DejaVu", "", font_path)
    pdf.add_font("DejaVu", "B", font_bold_path)
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()

    pdf.set_font("DejaVu", "B", 14)
    pdf.set_text_color(30, 41, 59)
    pdf.cell(0, 10, "Dzīvojamās telpas īres līguma juridiskā atskaite",
             new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Izveides datums: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}",
             new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)

    for line in analysis_text.split("\n"):
        raw = line.strip()
        if not raw:
            pdf.ln(3)
            continue

        if raw.startswith("#"):
            pdf.ln(4)
            pdf.set_font("DejaVu", "B", 11)

            if "Kritiski" in raw or "🔴" in raw:
                pdf.set_text_color(220, 38, 38)
            elif "Brīdinājumi" in raw or "🟡" in raw:
                pdf.set_text_color(217, 119, 6)
            elif "Ieteikumi" in raw or "🟢" in raw:
                pdf.set_text_color(22, 163, 74)
            else:
                pdf.set_text_color(30, 41, 59)

            title_text = strip_emoji(clean_markdown(raw.lstrip("#"))).strip()
            pdf.multi_cell(0, 7, title_text, new_x="LMARGIN", new_y="NEXT")

            pdf.set_font("DejaVu", "", 9.5)
            pdf.set_text_color(40, 40, 40)
        else:
            pdf.set_font("DejaVu", "", 9.5)
            pdf.set_text_color(40, 40, 40)
            body = strip_emoji(clean_markdown(raw))
            # Aizvieto sarakstu zvaigznītes ar punktiņu
            body = re.sub(r"^[\*\-]\s+", "• ", body)
            pdf.multi_cell(0, 5.5, body, new_x="LMARGIN", new_y="NEXT")

    return io.BytesIO(bytes(pdf.output()))


# ---------------------------------------------------------------
# Poga līguma analīzei
# ---------------------------------------------------------------
analyze_button = st.button("🔍 Analizēt līgumu", type="primary")

if analyze_button:
    if not api_key:
        st.error("⚠️ Lūdzu, ievadi savu Gemini API atslēgu sānu joslā!")
    elif uploaded_file is None:
        st.error("⚠️ Lūdzu, augšupielādē līguma failu!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(
                'gemini-3.8-flash',
                generation_config=genai.GenerationConfig(
                    temperature=0,
                    top_p=1,
                    top_k=1
                )
            )

            file_extension = uploaded_file.name.split('.')[-1].lower()

            with st.spinner("🔄 Notiek līguma analīze un risku izvērtēšana..."):
                if file_extension == 'docx':
                    bytes_data = uploaded_file.getvalue()
                    target_stream = io.BytesIO(bytes_data)
                    doc = docx.Document(target_stream)
                    contract_text = "\n".join(
                        [p.text for p in doc.paragraphs if p.text.strip()]
                    )

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

                st.markdown("### 📥 Lejupielādēt analīzes atskaiti:")
                col1, col2 = st.columns(2)

                with col1:
                    try:
                        docx_file = generate_pretty_docx(full_response_text)
                        st.download_button(
                            label="📄 Lejupielādēt Word (.docx)",
                            data=docx_file,
                            file_name=f"Ires_Liguma_Analize_{datetime.date.today()}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.warning(f"Word failu neizdevās izveidot: {e}")

                with col2:
                    try:
                        pdf_file = generate_pretty_pdf(full_response_text)
                        st.download_button(
                            label="📄 Lejupielādēt PDF (.pdf)",
                            data=pdf_file,
                            file_name=f"Ires_Liguma_Analize_{datetime.date.today()}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.warning(f"PDF failu neizdevās izveidot: {e}")

        except Exception as e:
            st.error(f"❌ Kļūda apstrādājot failu: {str(e)}")
