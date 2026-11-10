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
- Pārskati līgumu punktu pa punktam, no sākuma līdz galam, un neizlaid nevienu būtisku punktu.
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
[Katru punktu numurē ar kodu K1, K2, K3... Formāts: "K1. Līguma punkta numurs (īss nosaukums): paskaidrojums. [[atsauce sur likumu]]"]

## 🟡 Brīdinājumi
[Katru punktu numurē ar kodu B1, B2, B3... Formāts: "B1. Līguma punkta numurs (īss nosaukums): paskaidrojums. [[atsauce uz likumu]]"]

## 🟢 Ieteikumi un labojumi
[OBLIGĀTI: katram kritiskajam riskam un katram brīdinājumam jābūt savam ieteikumam ar to pašu kodu (K1, K2... B1, B2...). Formāts: "K1. Līguma punkta numurs: ko mainīt vai pievienot. Piedāvātais formulējums: „...”". Ja punkts līgumā trūkst, ieteikumā dod gatavu punkta tekstu, ko pievienot. Neviens kods nedrīkst palikt bez ieteikuma.]
"""

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
    for marker in ("Tiesību akta pase", "Saistītie dokumenti"):
        idx = text.find(marker)
        if idx > 2000:
            text = text[:idx]
            break
    if len(text) < 5000:
        raise ValueError("Likuma teksts izskatās nepilnīgs")
    return text

def load_laws():
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
        return "\n\n".join(parts), f"likumi.lv (automātiski)"

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
Tu esi ļoti rūpīgs latviešu valodas korektors un līgumu redaktors. Tev ir dots īres līgums. Tavs uzdevums ir atrast VISAS neitrālās rediģēšanas un neuzmanības kļūdas.
FORMĀTS: Sāc ar virsraksti ## ⚪ Neitrālas piezīmes (korektūra) un numurē N1, N2...
"""

def build_prompt(base=None):
    laws_text, source = load_laws()
    prompt = base or SYSTEM_PROMPT
    if laws_text:
        prompt += "\n\nSPĒKĀ ESOŠIE LIKUMI:\n" + laws_text
    return prompt, source

LAW_LINK_INFO = {
    "DTIL": ("Dzīvojamo telpu īres likuma", "https://likumi.lv/ta/id/322216"),
    "CL": ("Civillikuma", "https://likumi.lv/ta/id/90220"),
}
CITATION_RE = re.compile(r"\[\[\s*(DTIL|CL)\s+(\d{1,4})\s*\]\]")
SOURCES_MD = "**Pilni likumu teksti likumi.lv:** [Dzīvojamo telpu īres likums](https://likumi.lv/ta/id/322216) · [Civillikums](https://likumi.lv/ta/id/90220)"

def _citation_to_link(match):
    code, number = match.group(1), match.group(2)
    name, base = LAW_LINK_INFO[code]
    return f"[{name} {number}. pants]({base}#p{number})"

def linkify(text):
    return CITATION_RE.sub(_citation_to_link, text)

def linkify_stream(chunks):
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

def strip_emoji(text):
    text = re.sub(r"[\U00010000-\U0010FFFF]", "", text)
    text = re.sub(r"[\u2600-\u27BF\uFE0F]", "", text)
    return text

def clean_markdown(text):
    text = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)", r"\1 (\2)", text)
    return text.replace("**", "").replace("__", "")

def generate_pretty_docx(analysis_text):
    doc = docx.Document()
    title = doc.add_heading("ĪRES LĪGUMA JURIDISKĀ ANALĪZE", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for line in analysis_text.split("\n"):
        line_str = line.strip()
        if not line_str:
            continue
        if line_str.startswith("#"):
            doc.add_heading(strip_emoji(clean_markdown(line_str.lstrip("#"))).strip(), level=1)
        else:
            doc.add_paragraph(clean_markdown(line_str))
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# PDF un citu palīgfunkciju integrācija paliek nemainīga...
