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
import requests
from bs4 import BeautifulSoup

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

OBLIGĀTĀ PĀRBAUDES SARAKSTS (pārbaudi KATRU no šīm tēmām vienā un tajā pašā secībā; ja līgumā tā trūkst vai ir problemātiska, atzīmē atbilstošā sadaļā):
1. Pušu dati un rekvizīti (vai ir paredzēta vieta vārdiem, personas kodiem, adresēm, bankas kontam, tālrunim un e-pastam), īpašuma tiesības un telpu apraksts
2. Līguma termiņš, pagarināšana un izbeigšanas kārtība (abu pušu tiesības)
3. Īres maksas apmērs, maksāšanas termiņš un maksas maiņas nosacījumi
4. Komunālie un apsaimniekošanas maksājumi (arī uzkrājumu fonds, rēķinu uzrādīšana)
5. Drošības depozīts (apmērs, ieturēšana, atmaksas termiņš)
6. Nokavējuma procenti un līgumsodi (vai nav "VAI" formulējumu, vai sodi ir samērīgi)
7. Remonts un uzturēšana (kas ir "sīkie remontdarbi", kas maksā par lielajiem)
8. Dzīvokļa apskates un izīrētāja piekļuve (brīdinājuma laiks, biežums)
9. Pieņemšanas-nodošanas akts un telpu stāvoklis
10. Apakšīre, personu skaits, dzīvnieki un citi lietošanas ierobežojumi
11. Īres tiesību nostiprināšana Zemesgrāmatā un aizsardzība īpašnieka maiņas gadījumā
12. Strīdu risināšana un piemērojamais likums
13. Tukši lauki finansiāli un juridiski būtiskos punktos (summas, termiņi, procenti), nenoteikti termini

TUKŠIE LAUKI (svarīgs noteikums):
- Tukšie lauki, kas jāaizpilda līguma parakstīšanas brīdī un kas ir personas vai identifikācijas dati, NAV kļūda, un tos NEDRĪKST atzīmēt nevienā sadaļā. Tie ir: vārdi, uzvārdi, personas kodi, deklarētās adreses, tālruņi, e-pasti, bankas konti, kadastra numuri, dokumentu numuri, paraksti, līguma noslēgšanas vieta un datums.
- Tukši vai neskaidri lauki, kas ir FINANSIĀLI vai JURIDISKI BŪTISKI, OBLIGĀTI jāatzīmē: īres maksas summa, drošības depozīta summa, komunālo maksājumu summas vai kārtība, maksājuma termiņš vai datums, nokavējuma procentu vai līgumsoda apmērs, līguma sākuma un beigu termiņš, uzteikuma termiņi, depozīta atmaksas termiņš. Tos atzīmē kā brīdinājumu (vai kā kritisku risku, ja tukšums rada nepamatotu finansiālu risku).

- TRŪKSTOŠA VIETA REKVIZĪTIEM: ja līgumā VISPĀR NAV paredzēta vieta vai punkts, kur norādīt svarīgus datus, tā ir kļūda, kas jāatzīmē kā brīdinājums. Tas attiecas uz: izīrētāja bankas konta (IBAN) un maksājuma saņēmēja datiem, kam jāveic īres maksa; pušu kontaktinformāciju (tālrunis, e-pasts) paziņojumu nosūtīšanai; pušu adresēm vai juridiskajām adresēm; personas kodiem vai reģistrācijas numuriem; kadastra apzīmējumu vai precīzu dzīvokļa adresi; parakstu vietu. Ja šāda vieta līgumā ir, bet tā ir tikai neaizpildīta, tā NAV kļūda.
- Atšķirība: tukšs lauks "Konta Nr. ________" ir kārtībā, bet līgums, kurā nav nevienas vietas konta numuram vai nav norādīts, kā un kur maksāt īres maksu, ir jāatzīmē.

DARBA METODE:
- Pārskati līgumu punktu pa punktam, no sākuma līdz beigām, un neizlaid nevienu būtisku punktu.
- Esi rūpīgs un kritisks. Labāk atzīmē vairāk potenciālu problēmu nekā mazāk.
- Katrā sadaļā punktus kārto tajā secībā, kādā tie ir līgumā (no pirmā līdz pēdējam), nevis pēc svarīguma. Tas nodrošina vienādu rezultātu atkārtotā analīzē.
- Katram punktam norādi līguma punkta numuru un vienkāršu paskaidrojumu, kāpēc tas ir problēma.

KLASIFIKĀCIJA (katrs punkts drīkst būt tikai VIENĀ sadaļā):

🔴 KRITISKS RISKS - punkts ir kritisks, ja tas atbilst vismaz vienam no šiem nosacījumiem:
A) Tas PĀRKĀPJ LIKUMU vai ir pretrunā ar likuma imperatīvajām normām (Dzīvojamo telpu īres likumu, Civillikumu), tāpēc var būt spēkā neesošs. Vienmēr norādi, kuru likuma pantu tas pārkāpj.
B) Tas rada NEPAMATOTUS FINANSIĀLUS ZAUDĒJUMUS vienai pusei, piemēram:
   - depozīta patvaļīga ieturēšana vai neatmaksāšana;
   - nesamērīgs līgumsods vai nokavējuma procenti, kas pārsniedz likumā pieļauto;
   - maksājumi, kas pēc likuma vai labas prakses nav jāsedz šai pusei (piemēram, kapitālā remonta vai uzkrājumu fonda maksājumi, ko sedz īrnieks);
   - atbildība par zaudējumiem bez vainas vai bez pierādījumiem;
   - vienpusēja īres maksas palielināšana bez likumā noteiktā pamata un kārtības.
Ja punkts nepārkāpj likumu un nerada nepamatotus finansiālus zaudējumus, tas NAV kritisks.

🟡 BRĪDINĀJUMS - punkts, kas nav līdz galam skaidrs vai kura līgumā trūkst, un kuru vajadzētu pievienot vai precizēt, lai nebūtu šaubu dažādās situācijās:
- neskaidri vai vispārīgi formulējumi (piemēram, "sīkie remontdarbi", "citi maksājumi", "iepriekš saskaņojot");
- trūkstoši termiņi, summas, procenti vai kārtība (piemēram, nav noteikts brīdinājuma laiks apskatei, depozīta atmaksas termiņš, rēķinu uzrādīšanas kārtība);
- situācijas, ko līgums neapskata (piemēram, īpašnieka maiņa, līguma pagarināšana, strīdu risināšana, apakšīre, dzīvnieki, bojājumi nelaimes gadījumā);
- neskaidrs vai pretrunīgs formulējums, kas var radīt dažādas interpretācijas (piemēram, "VAI" starp diviem sodiem).
Brīdinājums NAV pārkāpums un nerada tiešus zaudējumus, bet atstāj vietu strīdam.

🟢 IETEIKUMS - konkrēts labojums vai papildinājums ar piedāvātu jaunu formulējumu.

Pirms atbildes pārbaudi, lai neviens punkts neatkārtotos vairākās sadaļās.
Ja kādā sadaļā tiešām nav problēmu, to godīgi pasaki, bet vispirms vēlreiz pārbaudi visu līgumu.

ATSAUCES UZ LIKUMIEM:
- Katram punktam sadaļās "Kritiski riski" un "Brīdinājumi" beigās pievieno juridisko pamatojumu, ja tāds ir sniegtajā likumu tekstā.
- Atsauci raksti TIKAI šādā formātā: [[DTIL 28]] (Dzīvojamo telpu īres likuma 28. pants) vai [[CL 2174]] (Civillikuma 2174. pants). Cipars ir tikai panta numurs.
- Ja punktam ir vairāki pamati, raksti vairākas atsauces pēc kārtas, piemēram: [[DTIL 28]] [[CL 2174]].
- Atsauci drīkst izmantot tikai tad, ja šis pants tiešām ir sniegtajā likumu tekstā. Neizdomā pantu numurus. Ja neesi pārliecināts, atsauci nerakstīt.
- Nerakstī interneta adreses, tās tiks pievienotas automātiski.

Izmanto tieši šādu virsrakstu struktūru:

## 📊 Kopējais novērtējums
[1-2 teikumos pasaki, vai šo līgumu drīkst parakstīt un kāds ir galvenais risks]

## 🔴 Kritiski riski
[Katru punktu numurē ar kodu K1, K2, K3... Formāts: "K1. Līguma punkta numurs (īss nosaukums): paskaidrojums. [[atsauce uz likumu]]"]

## 🟡 Brīdinājumi
[Katru punktu numurē ar kodu B1, B2, B3... Formāts: "B1. Līguma punkta numurs (īss nosaukums): paskaidrojums. [[atsauce uz likumu]]"]

## 🟢 Ieteikumi un labojumi
[OBLIGĀTI: katram kritiskajam riskam un katram brīdinājumam jābūt savam ieteikumam ar to pašu kodu (K1, K2... B1, B2...). Formāts: "K1. Līguma punkta numurs: ko mainīt vai pievienot. Piedāvātais formulējums: „...”". Ja punkts līgumā trūkst, ieteikumā dod gatavu punkta tekstu, ko pievienot. Neviens kods nedrīkst palikt bez ieteikuma.]

PĒDĒJĀ PĀRBAUDE PIRMS ATBILDES: saskaiti kodus sadaļās "Kritiski riski" un "Brīdinājumi" un pārliecinies, ka sadaļā "Ieteikumi un labojumi" ir ieteikums katram no tiem. Ja kāds kods trūkst, pievieno ieteikumu pirms atbildes beigām.
"""


# ---------------------------------------------------------------
# Likumu teksts: automātiski no likumi.lv (atjaunojas reizi 24 stundās)
# Ja neizdodas - izmanto rezerves failu likumi.txt (ja tāds ir repozitorijā)
# ---------------------------------------------------------------
# Šeit vari pievienot vēl likumus: (nosaukums, adrese ar "-print" beigās, panti vai None)
# Ja panti ir norādīti, no lielā likuma tiek paņemti tikai šie pantu diapazoni.
LAW_SOURCES = [
    ("DZĪVOJAMO TELPU ĪRES LIKUMS", "https://likumi.lv/ta/id/322216-print", None),
    (
        "CIVILLIKUMS, CETURTĀ DAĻA (izvilkums: līgumsods, procenti, īre un noma)",
        "https://likumi.lv/ta/id/90220-print",
        [(1716, 1726), (1759, 1765), (2100, 2180)],
    ),
]

ARTICLE_START = re.compile(r"^(\d{4})(?:\.\d+)?\.?\s+(?!gad)")


def filter_articles(text, ranges):
    """Atstāj tikai tos pantus, kuru numuri ietilpst norādītajos diapazonos."""
    kept = []
    include = False
    for line in text.split("\n"):
        m = ARTICLE_START.match(line)
        if m:
            num = int(m.group(1))
            include = any(lo <= num <= hi for lo, hi in ranges)
        if include:
            kept.append(line)
    return "\n".join(kept)


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_law(url):
    headers = {"User-Agent": "Mozilla/5.0 (compatible; IresLigumuSkeneris/1.0)"}
    r = requests.get(url, headers=headers, timeout=60)
    r.raise_for_status()
    r.encoding = "utf-8"
    soup = BeautifulSoup(r.text, "html.parser")
    for tag in soup(["script", "style", "nav", "header", "footer", "noscript"]):
        tag.decompose()
    text = soup.get_text("\n")
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    text = "\n".join(lines)
    # Noņem lapas apakšas informāciju
    for marker in ("Tiesību akta pase", "Saistītie dokumenti"):
        idx = text.find(marker)
        if idx > 2000:
            text = text[:idx]
            break
    if len(text) < 5000:
        raise ValueError("Likuma teksts izskatās nepilnīgs")
    return text


def load_laws():
    """Atgriež (likumu teksts, avota apraksts)."""
    parts = []
    for name, url, ranges in LAW_SOURCES:
        try:
            text = fetch_law(url)
            if ranges:
                text = filter_articles(text, ranges)
                if len(text) < 2000:
                    raise ValueError("Nav atrasti vajadzīgie panti")
            parts.append(f"=== {name} ===\n{text}")
        except Exception:
            pass
    if parts:
        return "\n\n".join(parts), f"likumi.lv (automātiski, {len(parts)} no {len(LAW_SOURCES)} likumiem)"

    laws_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "likumi.txt")
    if os.path.exists(laws_path):
        try:
            with open(laws_path, "r", encoding="utf-8") as f:
                txt = f.read().strip()
            if txt:
                return txt, "rezerves fails likumi.txt"
        except Exception:
            pass
    return "", "nav pieejams"


PROOFREAD_PROMPT = """
Tu esi ļoti rūpīgs latviešu valodas korektors un līgumu redaktors. Tev ir dots īres līgums. Tavs uzdevums ir atrast VISAS neitrālās rediģēšanas un neuzmanības kļūdas, pat vismazākās. Tu nevērtē līguma juridisko saturu un neanalizē riskus.

MEKLĒ:
- gramatikas, pareizrakstības un drukas kļūdas (arī garumzīmju un mīkstinājuma zīmju kļūdas);
- lieki vai trūkstoši vārdi, dubultas atstarpes, nepareizas vai trūkstošas pieturzīmes;
- nekonsekventi termini vai pušu apzīmējumi (piemēram, "Izīrētājs" un "Iznomātājs", "Īrnieks" un "Nomnieks" vienā līgumā);
- nepareizi punktu numuri, izlaisti numuri vai atsauces uz neesošiem līguma punktiem;
- pretrunas starp summām, datumiem vai termiņiem (arī starp cipariem un vārdiem);
- neveikli vai grūti saprotami teikumi bez juridiskas ietekmes.

NEATZĪMĒ:
- tukšus laukus, kas jāaizpilda parakstīšanas brīdī (vārdi, uzvārdi, personas kodi, adreses, tālruņi, e-pasti, bankas konti, kadastra numuri, paraksti, vieta un datums);
- juridiskus riskus, likuma pārkāpumus vai finansiālus jautājumus (tos izvērtē cits modulis).

Īpašnieks pats izlems, ko labot un ko atstāt, tāpēc atzīmē visu.

FORMĀTS:
- Sāc atbildi tieši ar virsrakstu: ## ⚪ Neitrālas piezīmes (korektūra)
- Katru piezīmi numurē N1, N2, N3... tajā secībā, kādā tās ir līgumā.
- Formāts: "N1. Līguma punkta numurs: kas ir nepareizi (kā tas ir uzrakstīts līgumā). Labojums: „pareizais variants”."
- Ja tiešām nekādu kļūdu nav, zem virsraksta raksti tikai: Korektūras kļūdu netika atrasts.
- Raksti latviešu valodā.
"""


def build_prompt(base=None):
    laws_text, source = load_laws()
    prompt = base or SYSTEM_PROMPT
    if laws_text:
        prompt += (
            "\n\nSPĒKĀ ESOŠIE LIKUMI (OFICIĀLAIS TEKSTS):\n"
            "Analīzē balsties uz šiem likumu tekstiem. Atsaucies uz pantiem tikai tad, ja tie ir "
            "šajā tekstā. Ja neesi pārliecināts par panta numuru, nenorādi to un neizdomā.\n\n"
            + laws_text
        )
    return prompt, source


# ---------------------------------------------------------------
# Atsauces uz likumi.lv: [[DTIL 28]] -> klikšķināma saite uz attiecīgo pantu
# ---------------------------------------------------------------
LAW_LINK_INFO = {
    "DTIL": ("Dzīvojamo telpu īres likuma", "https://likumi.lv/ta/id/322216"),
    "CL": ("Civillikuma", "https://likumi.lv/ta/id/90220"),
}
CITATION_RE = re.compile(r"\[\[\s*(DTIL|CL)\s+(\d{1,4})\s*\]\]")

SOURCES_MD = (
    "**Pilni likumu teksti likumi.lv:** "
    "[Dzīvojamo telpu īres likums](https://likumi.lv/ta/id/322216) · "
    "[Civillikuma ceturtā daļa](https://likumi.lv/ta/id/90220)"
)


def _citation_to_link(match):
    code, number = match.group(1), match.group(2)
    name, base = LAW_LINK_INFO[code]
    return f"[{name} {number}. pants]({base}#p{number})"


def linkify(text):
    return CITATION_RE.sub(_citation_to_link, text)


def linkify_stream(chunks):
    """Straumē tekstu un uz vietas pārvērš atsauces [[...]] par saitēm."""
    buf = ""
    for chunk in chunks:
        buf += chunk
        buf = linkify(buf)
        hold_from = buf.rfind("[[")
        if hold_from == -1 and buf.endswith("["):
            hold_from = len(buf) - 1
        if hold_from != -1 and "]]" not in buf[hold_from:]:
            out, buf = buf[:hold_from], buf[hold_from:]
        else:
            out, buf = buf, ""
        if out:
            yield out
    if buf:
        yield linkify(buf)


# ---------------------------------------------------------------
# Palīgfunkcijas teksta tīrīšanai
# ---------------------------------------------------------------
def strip_emoji(text):
    """Noņem emocijzīmes (DejaVu fonts tās neatbalsta PDF failā)."""
    text = re.sub(r"[\U00010000-\U0010FFFF]", "", text)
    text = re.sub(r"[\u2600-\u27BF\uFE0F]", "", text)
    return text


def clean_markdown(text):
    """Noņem markdown zīmes; saites pārvērš par 'teksts (adrese)' Word un PDF failiem."""
    text = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)", r"\1 (\2)", text)
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
            elif "Neitrāl" in line_str or "⚪" in line_str:
                color = RGBColor(100, 116, 139)
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
            elif "Neitrāl" in raw or "⚪" in raw:
                pdf.set_text_color(100, 116, 139)
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


def build_downloads(report_text):
    """Izveido Word un PDF failus (baitos), lai tos varētu lejupielādēt vairākkārt."""
    result = {"docx": None, "pdf": None, "docx_err": None, "pdf_err": None}
    try:
        result["docx"] = generate_pretty_docx(report_text).getvalue()
    except Exception as e:
        result["docx_err"] = str(e)
    try:
        result["pdf"] = generate_pretty_pdf(report_text).getvalue()
    except Exception as e:
        result["pdf_err"] = str(e)
    return result


def render_downloads(result):
    st.success("✅ Analīze veiksmīgi pabeigta!")
    st.markdown("---")
    st.markdown("### 📥 Lejupielādēt analīzes atskaiti:")
    col1, col2 = st.columns(2)
    today = datetime.date.today()

    with col1:
        if result["docx"]:
            st.download_button(
                label="📄 Lejupielādēt Word (.docx)",
                data=result["docx"],
                file_name=f"Ires_Liguma_Analize_{today}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="dl_docx"
            )
        else:
            st.warning(f"Word failu neizdevās izveidot: {result['docx_err']}")

    with col2:
        if result["pdf"]:
            st.download_button(
                label="📄 Lejupielādēt PDF (.pdf)",
                data=result["pdf"],
                file_name=f"Ires_Liguma_Analize_{today}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="dl_pdf"
            )
        else:
            st.warning(f"PDF failu neizdevās izveidot: {result['pdf_err']}")


# ---------------------------------------------------------------
# Analīzes izpilde
# ---------------------------------------------------------------
if analyze_button:
    # Jauna analīze: izdzēš iepriekšējo saglabāto rezultātu
    st.session_state.pop("result", None)

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
                active_prompt, laws_source = build_prompt()
                if file_extension == 'docx':
                    bytes_data = uploaded_file.getvalue()
                    target_stream = io.BytesIO(bytes_data)
                    doc = docx.Document(target_stream)
                    contract_text = "\n".join(
                        [p.text for p in doc.paragraphs if p.text.strip()]
                    )

                    full_prompt = f"{active_prompt}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{contract_text}"
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

                    response = model.generate_content([active_prompt, file_part], stream=True)

            st.caption(f"📚 Likumu avots: {laws_source}")
            st.markdown("### 📋 Analīzes rezultāts reāllaikā:")

            def stream_and_collect():
                for chunk in response:
                    if hasattr(chunk, 'text') and chunk.text:
                        yield chunk.text

            full_response_text = st.write_stream(linkify_stream(stream_and_collect()))

            # Otrā pārbaude: atsevišķa korektūra (bez likumu teksta, tāpēc lētāka)
            proof_text = ""
            if full_response_text:
                try:
                    st.markdown("---")
                    with st.spinner("✏️ Notiek korektūras pārbaude..."):
                        if file_extension == 'docx':
                            proof_input = f"{PROOFREAD_PROMPT}\n\nLĪGUMA TEKSTS:\n{contract_text}"
                        else:
                            proof_input = [PROOFREAD_PROMPT, file_part]
                        proof_response = model.generate_content(proof_input, stream=True)

                    def stream_proof():
                        for chunk in proof_response:
                            if hasattr(chunk, 'text') and chunk.text:
                                yield chunk.text

                    proof_text = st.write_stream(stream_proof())
                except Exception as e:
                    st.warning(f"Korektūras pārbaudi neizdevās veikt: {e}")

            if full_response_text:
                analysis_all = full_response_text
                if proof_text:
                    analysis_all += "\n\n" + proof_text
                st.markdown(SOURCES_MD)
                report_text = analysis_all + "\n\n## Izmantotie avoti\n" + SOURCES_MD
                result = build_downloads(report_text)
                result.update({
                    "analysis": analysis_all,
                    "laws_source": laws_source,
                })
                # Saglabā, lai analīze nepazustu pēc lejupielādes
                st.session_state["result"] = result
                render_downloads(result)

        except Exception as e:
            st.error(f"❌ Kļūda apstrādājot failu: {str(e)}")

elif "result" in st.session_state:
    # Lapa pārlādējās (piem., nospiesta lejupielādes poga): parāda saglabāto analīzi
    saved = st.session_state["result"]
    st.caption(f"📚 Likumu avots: {saved['laws_source']}")
    st.markdown("### 📋 Analīzes rezultāts:")
    st.markdown(saved["analysis"])
    st.markdown(SOURCES_MD)
    render_downloads(saved)
