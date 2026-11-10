import streamlit as st
import google.generativeai as genai
from google.generativeai import caching
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import io
import re
import json
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

# ---------------------------------------------------------------
# Interfeisa valoda (tikai lietotāja saskarne; analīze vienmēr ir latviski)
# ---------------------------------------------------------------
TR = {
    "Latviešu": {
        "title": "🏠 Dzīvokļu Īres Līgumu AI Skeneris",
        "subtitle": "Augšupielādē savu līgumu un saņem tūlītēju reāllaika straumēšanas analīzi.",
        "sidebar_header": "🔑 Piekļuves iestatījumi",
        "api_label": "Ievadi Gemini API atslēgu",
        "uploader": "Izvēlies vai ievelc līguma failu (.docx, .pdf, .png, .jpg):",
        "analyze_btn": "🔍 Analizēt līgumu",
        "err_no_key": "⚠️ Lūdzu, ievadi savu Gemini API atslēgu sānu joslā!",
        "err_no_file": "⚠️ Lūdzu, augšupielādē līguma failu vai ielīmē līguma tekstu!",
        "spinner_analysis": "🔄 Notiek līguma analīze un risku izvērtēšana...",
        "law_source": "📚 Likumu avots",
        "result_live": "### 📋 Analīzes rezultāts reāllaikā:",
        "result_saved": "### 📋 Analīzes rezultāts:",
        "proof_spinner": "✏️ Notiek korektūras pārbaude...",
        "proof_fail": "Korektūras pārbaudi neizdevās veikt",
        "success": "✅ Analīze veiksmīgi pabeigta!",
        "dl_header": "### 📥 Lejupielādēt analīzes atskaiti:",
        "dl_word": "📄 Lejupielādēt Word (.docx)",
        "dl_pdf": "📄 Lejupielādēt PDF (.pdf)",
        "word_fail": "Word failu neizdevās izveidot",
        "pdf_fail": "PDF failu neizdevās izveidot",
        "err_process": "❌ Kļūda apstrādājot failu",
    },
    "English": {
        "title": "🏠 Apartment Lease AI Scanner",
        "subtitle": "Upload your contract and get an instant real-time analysis.",
        "sidebar_header": "🔑 Access settings",
        "api_label": "Enter your Gemini API key",
        "uploader": "Choose or drag the contract file (.docx, .pdf, .png, .jpg):",
        "analyze_btn": "🔍 Analyze contract",
        "err_no_key": "⚠️ Please enter your Gemini API key in the sidebar!",
        "err_no_file": "⚠️ Please upload a contract file or paste the contract text!",
        "spinner_analysis": "🔄 Analyzing the contract and assessing risks...",
        "law_source": "📚 Law source",
        "result_live": "### 📋 Analysis result (live):",
        "result_saved": "### 📋 Analysis result:",
        "proof_spinner": "✏️ Proofreading check in progress...",
        "proof_fail": "Proofreading check failed",
        "success": "✅ Analysis completed successfully!",
        "dl_header": "### 📥 Download the analysis report:",
        "dl_word": "📄 Download Word (.docx)",
        "dl_pdf": "📄 Download PDF (.pdf)",
        "word_fail": "Could not create the Word file",
        "pdf_fail": "Could not create the PDF file",
        "err_process": "❌ Error processing the file",
    },
    "Русский": {
        "title": "🏠 ИИ-сканер договоров аренды квартир",
        "subtitle": "Загрузите договор и получите мгновенный анализ в реальном времени.",
        "sidebar_header": "🔑 Настройки доступа",
        "api_label": "Введите API-ключ Gemini",
        "uploader": "Выберите или перетащите файл договора (.docx, .pdf, .png, .jpg):",
        "analyze_btn": "🔍 Проанализировать договор",
        "err_no_key": "⚠️ Пожалуйста, введите API-ключ Gemini в боковой панели!",
        "err_no_file": "⚠️ Пожалуйста, загрузите файл договора или вставьте текст договора!",
        "spinner_analysis": "🔄 Идёт анализ договора и оценка рисков...",
        "law_source": "📚 Источник законов",
        "result_live": "### 📋 Результат анализа (в реальном времени):",
        "result_saved": "### 📋 Результат анализа:",
        "proof_spinner": "✏️ Выполняется корректорская проверка...",
        "proof_fail": "Не удалось выполнить корректорскую проверку",
        "success": "✅ Анализ успешно завершён!",
        "dl_header": "### 📥 Скачать отчёт об анализе:",
        "dl_word": "📄 Скачать Word (.docx)",
        "dl_pdf": "📄 Скачать PDF (.pdf)",
        "word_fail": "Не удалось создать файл Word",
        "pdf_fail": "Не удалось создать файл PDF",
        "err_process": "❌ Ошибка при обработке файла",
    },
}

TR_EXTRA = {
    "Latviešu": {
        "report_lang_name": "latviešu",
        "paste_label": "Vai ielīmē līguma tekstu šeit (ja ielīmēts teksts, fails netiek izmantots):",
        "h_overall": "Kopējais novērtējums", "h_crit": "Kritiski riski", "h_warn": "Brīdinājumi",
        "h_recs": "Ieteikumi un labojumi", "h_border": "Apsveramie punkti (robežgadījumi)", "h_neutral": "Neitrālas piezīmes (korektūra)",
        "no_proof_errors": "Korektūras kļūdu netika atrasts.",
        "label_proposed": "Piedāvātais formulējums", "label_translation": "Tulkojums", "label_fix": "Labojums",
        "sources_head": "Izmantotie avoti", "sources_label": "Pilni likumu teksti likumi.lv:",
        "law1_name": "Dzīvojamo telpu īres likums", "law2_name": "Civillikuma ceturtā daļa",
        "cite_dtil": "Dzīvojamo telpu īres likuma {n}. pants", "cite_cl": "Civillikuma {n}. pants",
        "doc_title": "ĪRES LĪGUMA JURIDISKĀ ANALĪZE", "doc_prepared": "Atskaite sagatavota",
        "pdf_header": "JURIDISKAIS ATZINUMS — ĪRES LĪGUMA ANALĪZE", "pdf_page": "Lappuse",
        "pdf_footer": "AI Juridiskais Asistents",
        "pdf_title": "Dzīvojamās telpas īres līguma juridiskā atskaite", "pdf_date": "Izveides datums",
        "up_browse": "Izvēlēties failu", "up_drag": "Ievelc failu šeit",
    },
    "English": {
        "report_lang_name": "angļu",
        "paste_label": "Or paste the contract text here (if text is pasted, the file is not used):",
        "h_overall": "Overall assessment", "h_crit": "Critical risks", "h_warn": "Warnings",
        "h_recs": "Recommendations and corrections", "h_border": "Points to consider (borderline cases)", "h_neutral": "Neutral notes (proofreading)",
        "no_proof_errors": "No proofreading errors were found.",
        "label_proposed": "Proposed wording", "label_translation": "Translation", "label_fix": "Correction",
        "sources_head": "Sources used", "sources_label": "Full law texts on likumi.lv:",
        "law1_name": "Residential Tenancy Law (Dzīvojamo telpu īres likums)",
        "law2_name": "Civil Law, Part Four (Civillikums)",
        "cite_dtil": "Residential Tenancy Law, Art. {n}", "cite_cl": "Civil Law, Art. {n}",
        "doc_title": "LEGAL ANALYSIS OF THE LEASE AGREEMENT", "doc_prepared": "Report prepared",
        "pdf_header": "LEGAL OPINION — LEASE AGREEMENT ANALYSIS", "pdf_page": "Page",
        "pdf_footer": "AI Legal Assistant",
        "pdf_title": "Legal report on the residential lease agreement", "pdf_date": "Date created",
        "up_browse": "Browse files", "up_drag": "Drag and drop file here",
    },
    "Русский": {
        "report_lang_name": "krievu",
        "paste_label": "Или вставьте текст договора сюда (если текст вставлен, файл не используется):",
        "h_overall": "Общая оценка", "h_crit": "Критические риски", "h_warn": "Предупреждения",
        "h_recs": "Рекомендации и исправления", "h_border": "Пункты для рассмотрения (пограничные случаи)", "h_neutral": "Нейтральные замечания (корректура)",
        "no_proof_errors": "Ошибок при корректуре не найдено.",
        "label_proposed": "Предлагаемая формулировка", "label_translation": "Перевод", "label_fix": "Исправление",
        "sources_head": "Использованные источники", "sources_label": "Полные тексты законов на likumi.lv:",
        "law1_name": "Закон о найме жилых помещений (Dzīvojamo telpu īres likums)",
        "law2_name": "Гражданский закон, часть четвёртая (Civillikums)",
        "cite_dtil": "Закон о найме жилых помещений, ст. {n}", "cite_cl": "Гражданский закон, ст. {n}",
        "doc_title": "ЮРИДИЧЕСКИЙ АНАЛИЗ ДОГОВОРА АРЕНДЫ", "doc_prepared": "Отчёт подготовлен",
        "pdf_header": "ЮРИДИЧЕСКОЕ ЗАКЛЮЧЕНИЕ — АНАЛИЗ ДОГОВОРА АРЕНДЫ", "pdf_page": "Страница",
        "pdf_footer": "ИИ-юридический помощник",
        "pdf_title": "Юридический отчёт по договору аренды жилого помещения", "pdf_date": "Дата создания",
        "up_browse": "Выбрать файл", "up_drag": "Перетащите файл сюда",
    },
}
for _l, _extra in TR_EXTRA.items():
    TR[_l].update(_extra)

# Papildu valodas: tulkojumu veic Gemini (kad ievadīta API atslēga)
EXTRA_LANGS = {
    "Deutsch": "German", "Español": "Spanish", "Français": "French", "Italiano": "Italian",
    "Português": "Portuguese", "Polski": "Polish", "Lietuvių": "Lithuanian", "Eesti": "Estonian",
    "Suomi": "Finnish", "Svenska": "Swedish", "Norsk": "Norwegian", "Dansk": "Danish",
    "Nederlands": "Dutch", "Українська": "Ukrainian", "Беларуская": "Belarusian", "Türkçe": "Turkish",
    "Čeština": "Czech", "Slovenčina": "Slovak", "Magyar": "Hungarian", "Română": "Romanian",
    "Български": "Bulgarian", "Ελληνικά": "Greek", "Српски": "Serbian", "Hrvatski": "Croatian",
    "Shqip": "Albanian", "Қазақша": "Kazakh", "Հայերեն": "Armenian", "ქართული": "Georgian",
    "中文": "Chinese", "日本語": "Japanese", "한국어": "Korean", "العربية": "Arabic",
    "עברית": "Hebrew", "हिन्दी": "Hindi", "اردو": "Urdu", "Tiếng Việt": "Vietnamese",
    "ไทย": "Thai", "Bahasa Indonesia": "Indonesian", "Filipino": "Filipino",
}
OTHER_LANG = "🌍 Cita valoda / Other / Другой…"


@st.cache_data(show_spinner=False)
def translate_ui(lang_name, _api_key):
    """Iztulko visus saskarnes tekstus izvēlētajā valodā (rezultāts tiek saglabāts kešatmiņā)."""
    genai.configure(api_key=_api_key)
    model = genai.GenerativeModel(
        'gemini-3.8-flash',
        generation_config=genai.GenerationConfig(temperature=0, response_mime_type="application/json"),
    )
    base = {k: v for k, v in TR["English"].items() if k != "report_lang_name"}
    prompt = (
        f"Translate the values of this JSON object into {lang_name}. They are user-interface texts of a web "
        "app that analyzes apartment lease contracts. Rules: keep every key unchanged; keep emojis, markdown "
        "prefixes such as '###', and placeholders such as {n} exactly as they are; keep 'likumi.lv', 'Gemini', "
        "'API', '.docx', '.pdf' and the Latvian names in parentheses unchanged; if a value is written in "
        "uppercase keep it uppercase; keep short button labels short. Return ONLY the JSON object.\n\n"
        + json.dumps(base, ensure_ascii=False)
    )
    data = json.loads(model.generate_content(prompt).text)
    out = dict(TR["English"])
    for k, v in data.items():
        if k in base and isinstance(v, str) and v.strip():
            out[k] = v
    out["report_lang_name"] = lang_name
    return out


lang_choice = st.sidebar.selectbox(
    "🌐 Valoda / Language / Язык",
    list(TR.keys()) + list(EXTRA_LANGS.keys()) + [OTHER_LANG],
)

if lang_choice in TR:
    T = TR[lang_choice]
else:
    if lang_choice == OTHER_LANG:
        lang_name = st.sidebar.text_input(
            "Language (any language in the world / jebkura pasaules valoda)", key="custom_lang"
        ).strip()
    else:
        lang_name = EXTRA_LANGS[lang_choice]

    T = dict(TR["English"])
    if lang_name:
        T["report_lang_name"] = lang_name
        _key = st.session_state.get("api_key_input", "")
        if _key:
            try:
                with st.spinner("🌐 Translating the interface..."):
                    T = translate_ui(lang_name, _key)
            except Exception as _e:
                st.sidebar.warning(f"Interface translation failed, showing English: {str(_e)[:120]}")
        else:
            st.sidebar.info(
                "🌐 Enter your Gemini API key below and the interface and the analysis "
                "will switch to this language."
            )


def _css_text(text):
    return text.replace("\\", "").replace('"', "'").replace("\n", " ")


# Faila augšupielādes pogas un uzraksta tulkojums: oriģinālais teksts + " / " + tulkojums
# (CSS; ja Streamlit maina struktūru, paliks tikai noklusējuma teksts)
if T["up_browse"] != TR["English"]["up_browse"]:
    st.markdown(
        "<style>"
        '[data-testid="stFileUploaderDropzone"] button::after { content: " / ' + _css_text(T["up_browse"]) + '"; white-space: pre; }'
        '[data-testid="stFileUploaderDropzoneInstructions"] > div > span::after { content: " / ' + _css_text(T["up_drag"]) + '"; }'
        "</style>",
        unsafe_allow_html=True,
    )

# Tēmas pārslēgšana: 🖥️ automātiski, ☀️ gaišā, 🌙 tumšā, ⚫ OLED (maksimāli melna).
# Pārslēdz pārlūkā (JavaScript), tāpēc Streamlit netiek pārlādēts un analīze netiek pārtraukta.
# Izvēle tiek atcerēta pārlūkā.
import streamlit.components.v1 as components

THEME_HTML = """
<style>
  body { margin: 0; font-family: sans-serif; }
  .row { display: flex; gap: 6px; }
  button { flex: 1; padding: 8px 0; font-size: 18px; border-radius: 8px; border: 1px solid #888;
           background: transparent; color: #999; cursor: pointer; }
  button.on { border-color: #ff4b4b; background: rgba(255,75,75,0.25); color: #fff; }
</style>
<div class="row">
  <button data-t="auto" title="Auto">🖥️</button>
  <button data-t="light" title="Light">☀️</button>
  <button data-t="dark" title="Dark">🌙</button>
  <button data-t="oled" title="OLED">⚫</button>
</div>
<script>
(function () {
  var P = window.parent, D = P.document;
  // [fons, sānu josla/lauki, teksts, apmales, saites]
  var THEMES = {
    dark:  ["#0e1117", "#1a1c24", "#fafafa", "#3a3d4a", "#6cb4ff"],
    light: ["#ffffff", "#f0f2f6", "#1f2430", "#d0d4dc", "#0b5cd5"],
    oled:  ["#000000", "#000000", "#e8e8e8", "#2a2a2a", "#6cb4ff"]
  };
  function css(name) {
    var c = THEMES[name], bg = c[0], bg2 = c[1], fg = c[2], bd = c[3], ln = c[4];
    var s =
      ".stApp, [data-testid='stHeader'] { background: " + bg + " !important; }" +
      "[data-testid='stSidebar'], [data-testid='stSidebar'] > div { background: " + bg2 + " !important; }" +
      ".stApp, .stApp p, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp small { color: " + fg + " !important; }" +
      ".stApp button p { color: inherit !important; }" +
      ".stApp a { color: " + ln + " !important; }" +
      "[data-testid='stAlert'] { background: rgba(128,128,128,0.18) !important; }" +
      "[data-testid='stAlert'] * { color: " + fg + " !important; }" +
      ".stTextInput input, .stTextArea textarea, [data-baseweb='select'] > div, [data-testid='stFileUploaderDropzone'] { background: " + bg2 + " !important; color: " + fg + " !important; }" +
      "[data-testid='stFileUploaderDropzoneInstructions'] * { color: " + fg + " !important; }" +
      "[data-testid='stExpander'] { border-color: " + bd + " !important; }" +
      "[data-testid='stExpander'] summary, [data-testid='stExpander'] summary * { color: " + fg + " !important; }" +
      "[data-baseweb='popover'] > div, [data-baseweb='popover'] [data-testid='stVerticalBlock'] { background: " + bg2 + " !important; }" +
      "[data-baseweb='popover'] p, [data-baseweb='popover'] li, [data-baseweb='popover'] small { color: " + fg + " !important; }";
    if (name === "oled") {
      s += ".stTextInput input, .stTextArea textarea, [data-baseweb='select'] > div, [data-testid='stFileUploaderDropzone'], [data-baseweb='popover'] > div { border: 1px solid " + bd + " !important; }";
    }
    return s;
  }
  function apply(t) {
    var s = D.getElementById("app-theme-style");
    if (!THEMES[t]) {
      if (s) s.remove();
    } else {
      if (!s) { s = D.createElement("style"); s.id = "app-theme-style"; D.head.appendChild(s); }
      s.textContent = css(t);
    }
    try { P.localStorage.setItem("appTheme", t); } catch (e) {}
    var bs = document.querySelectorAll("button");
    for (var i = 0; i < bs.length; i++) { bs[i].classList.toggle("on", bs[i].getAttribute("data-t") === t); }
  }
  var saved = "auto";
  try { saved = P.localStorage.getItem("appTheme") || "auto"; } catch (e) {}
  apply(saved);
  var btns = document.querySelectorAll("button");
  for (var i = 0; i < btns.length; i++) {
    btns[i].addEventListener("click", function () { apply(this.getAttribute("data-t")); });
  }
})();
</script>
"""

with st.sidebar:
    components.html(THEME_HTML, height=50)
st.title(T["title"])
st.markdown(T["subtitle"])

# Sānu josla
st.sidebar.header(T["sidebar_header"])
api_key = st.sidebar.text_input(T["api_label"], type="password", key="api_key_input")

uploaded_file = st.file_uploader(
    T["uploader"],
    type=["docx", "pdf", "png", "jpg", "jpeg"]
)

pasted_text = st.text_area(T["paste_label"], height=200, key="pasted_text")

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
- NEGRUPĒ: katru atsevišķu problēmu, katru atsevišķu tukšo finansiālo lauku (piemēram, īres summa, depozīta summa, maksājuma datums, termiņš) un katru atsevišķu līguma punktu atzīmē ATSEVIŠĶI ar savu kodu. Nekad nesavieno vairākus punktus vienā ierakstā.
- Ja šaubies, vai punkts ir problēma, NEIZLAID to: liec to sadaļā "Apsveramie punkti".
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
   - vienpusēja īres maksas palielināšana bez likumā noteiktā pamata un kārtības;
   - aizsardzības trūkums īpašnieka maiņas gadījumā: ja līgumā NAV paredzēta īres tiesību nostiprināšana Zemesgrāmatā, jaunais īpašnieks var prasīt atbrīvot telpas, un īrnieks zaudē mājokli un ar to saistītos izdevumus (pārcelšanās, depozīts, iepriekš samaksātā īre). Tas vienmēr jāatzīmē kā KRITISKS risks.
Ja punkts nepārkāpj likumu un nerada nepamatotus finansiālus zaudējumus, tas NAV kritisks.

🟡 BRĪDINĀJUMS - punkts, kas nav līdz galam skaidrs vai kura līgumā trūkst, un kuru vajadzētu pievienot vai precizēt, lai nebūtu šaubu dažādās situācijās:
- neskaidri vai vispārīgi formulējumi (piemēram, "sīkie remontdarbi", "citi maksājumi", "iepriekš saskaņojot");
- trūkstoši termiņi, summas, procenti vai kārtība (piemēram, nav noteikts brīdinājuma laiks apskatei, depozīta atmaksas termiņš, rēķinu uzrādīšanas kārtība);
- situācijas, ko līgums neapskata (piemēram, līguma pagarināšana, strīdu risināšana, apakšīre, dzīvnieki, bojājumi nelaimes gadījumā);
- neskaidrs vai pretrunīgs formulējums, kas var radīt dažādas interpretācijas (piemēram, "VAI" starp diviem sodiem).
Brīdinājums NAV pārkāpums un nerada tiešus zaudējumus, bet atstāj vietu strīdam.

🔵 APSVERAMS PUNKTS (robežgadījums) - punkts, kas ir līgumā, ir likumīgs un nerada acīmredzamus zaudējumus, bet kuru īpašnieks varētu vēlēties izvērtēt vai mainīt, jo tas ir mazāk izdevīgs vienai pusei, nestandarta vai atkarīgs no pušu vēlmēm. Īpašnieks pats izlems, vai tas jālabo.
Šeit liec VISU, par ko šaubies, vai tā ir problēma. Piemēri:
- drošības depozīta atmaksas termiņš, kas ir garāks par 10 darba dienām (piemēram, "1 mēneša laikā"): OBLIGĀTI šeit;
- garš vai vienpusējs uzteikuma termiņš, stingri vai nestandarta nosacījumi, kas ir likumīgi, bet nav līdzsvaroti;
- noteikumi, kas ir skaidri, bet var būt neizdevīgi (piemēram, depozīta apmērs, maksājumu termiņi, ierobežojumi viesiem vai remontiem);
- jebkas, kas nav ne skaidra kļūda, ne skaidri kārtībā.
Noteikums: ja šaubies, vai punkts ir problēma, NEIZLAID to, bet liec šajā sadaļā. Ja punkts ir skaidri nelikumīgs vai rada zaudējumus, tas ir kritisks; ja trūkst vai ir neskaidrs, tas ir brīdinājums; pārējos robežgadījumus liec šeit.

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

## 🔵 Apsveramie punkti (robežgadījumi)
[Katru punktu numurē ar kodu A1, A2, A3... Formāts: "A1. Līguma punkta numurs (īss nosaukums): kas tieši ir šaubīgs un kāpēc to varētu apsvērt. Ja vēlaties mainīt: piedāvātais formulējums: „...”". Atsauci uz likumu pievieno tikai, ja tā ir. Šeit ir arī punkti, ko īpašnieks var apzināti atstāt nemainītus.]

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

# Panta sākums: rinda "1716." (numurs atsevišķā rindā, jo likumi.lv to izceļ treknrakstā)
# vai "1716. Teksts..." Atsevišķi 4 ciparu skaitļi bez punkta (rindu numuri) netiek ņemti vērā.
ARTICLE_START = re.compile(r"^(\d{4})\.(?:\d{1,2})?$|^(\d{4})\.\s+(?!gad)\S")


def filter_articles(text, ranges):
    """Atstāj tikai tos pantus, kuru numuri ietilpst norādītajos diapazonos."""
    kept = []
    include = False
    for line in text.split("\n"):
        m = ARTICLE_START.match(line)
        if m:
            num = int(m.group(1) or m.group(2))
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
    errors = []
    for name, url, ranges in LAW_SOURCES:
        try:
            text = fetch_law(url)
            if ranges:
                text = filter_articles(text, ranges)
                if len(text) < 2000:
                    raise ValueError("Nav atrasti vajadzīgie panti")
            parts.append(f"=== {name} ===\n{text}")
        except Exception as e:
            errors.append(f"{name.split(',')[0].split('(')[0].strip()}: {str(e)[:80]}")
    if parts:
        info = f"likumi.lv (automātiski, {len(parts)} no {len(LAW_SOURCES)} likumiem)"
        if errors:
            info += " | neizdevās: " + "; ".join(errors)
        return "\n\n".join(parts), info

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


def lang_directive():
    rl = T["report_lang_name"]
    return (
        "Raksti vienkāršā, cilvēcīgā un viegli saprotamā valodā.\n\n"
        "VALODAS NOTEIKUMI (svarīgi):\n"
        f"- Visu analīzi (virsrakstus, paskaidrojumus, riskus, brīdinājumus un ieteikumu skaidrojumus) raksti {rl} valodā.\n"
        "- Nosaki, kādā valodā ir uzrakstīts līgums (latviešu, krievu vai angļu).\n"
        "- Citātus no līguma raksti oriģinālvalodā.\n"
        "- Konkrētus formulējumus, ko jāieraksta līgumā (piedāvātie formulējumi un jaunie punkti), VIENMĒR raksti LĪGUMA valodā pēdiņās. "
        f"Ja līguma valoda NAV {rl}, tad katram šādam formulējumam uzreiz zem tā pievieno tulkojumu {rl} valodā ar atzīmi \"{T['label_translation']}:\". "
        f"Ja līguma valoda ir {rl}, tulkojumu nepievieno.\n"
        "- Atsauces uz likumiem raksti tikai formātā [[DTIL 28]] vai [[CL 2174]], neatkarīgi no valodas.\n"
        "- Sadaļu virsrakstus raksti tieši tā, kā norādīts struktūrā zemāk."
    )


def localize_prompt(text):
    text = text.replace("Raksti vienkāršā, cilvēcīgā un viegli saprotamā latviešu valodā.", lang_directive())
    text = text.replace("## 📊 Kopējais novērtējums", "## 📊 " + T["h_overall"])
    text = text.replace("## 🔴 Kritiski riski", "## 🔴 " + T["h_crit"])
    text = text.replace("## 🟡 Brīdinājumi", "## 🟡 " + T["h_warn"])
    text = text.replace("## 🟢 Ieteikumi un labojumi", "## 🟢 " + T["h_recs"])
    text = text.replace("## 🔵 Apsveramie punkti (robežgadījumi)", "## 🔵 " + T["h_border"])
    text = text.replace("Piedāvātais formulējums", T["label_proposed"])
    return text


def build_proof_prompt():
    t = PROOFREAD_PROMPT
    rl = T["report_lang_name"]
    t = t.replace("Tu esi ļoti rūpīgs latviešu valodas korektors un līgumu redaktors.",
                  "Tu esi ļoti rūpīgs korektors un līgumu redaktors, kas pārzina līguma valodu (latviešu, krievu vai angļu).")
    t = t.replace("## ⚪ Neitrālas piezīmes (korektūra)", "## ⚪ " + T["h_neutral"])
    t = t.replace("Korektūras kļūdu netika atrasts.", T["no_proof_errors"])
    t = t.replace("Labojums:", T["label_fix"] + ":")
    t = t.replace("- Raksti latviešu valodā.",
                  f"- Skaidrojumus (kas ir nepareizi) raksti {rl} valodā. Citātu no līguma un pareizo variantu raksti līguma oriģinālvalodā. "
                  "Kļūdas meklē līguma valodas noteikumu ietvaros.")
    return t


def build_prompt(base=None):
    laws_text, source = load_laws()
    prompt = localize_prompt(base or SYSTEM_PROMPT)
    if laws_text:
        prompt += (
            "\n\nSPĒKĀ ESOŠIE LIKUMI (OFICIĀLAIS TEKSTS):\n"
            "Analīzē balsties uz šiem likumu tekstiem. Atsaucies uz pantiem tikai tad, ja tie ir "
            "šajā tekstā. Ja neesi pārliecināts par panta numuru, nenorādi to un neizdomā.\n\n"
            + laws_text
        )
    return prompt, source
    # ---------------------------------------------------------------
# Gemini Context Caching: SYSTEM_PROMPT + likumu teksts (TTL 24 stundas)
# ---------------------------------------------------------------
CACHE_MODEL = "models/gemini-3.8-flash"  # modelim jāatbalsta Context Caching
CACHE_TTL = datetime.timedelta(hours=24)


@st.cache_resource(ttl=datetime.timedelta(hours=23, minutes=30), show_spinner=False)
def create_prompt_cache(api_key, system_text, laws_block):
    """Izveido Gemini kešatmiņu. Streamlit to neveido atkārtoti 23,5 h laikā (Gemini kešatmiņas TTL ir 24 h)."""
    genai.configure(api_key=api_key)
    return caching.CachedContent.create(
        model=CACHE_MODEL,
        display_name="ires-liguma-skeneris",
        system_instruction=system_text,
        contents=[laws_block],
        ttl=CACHE_TTL,
    )


def get_cached_model(api_key, generation_config):
    """Atgriež (modelis ar kešatmiņu, likumu avots) vai (None, avots), ja kešatmiņu nevar izmantot."""
    laws_text, source = load_laws()
    if not laws_text:
        return None, source
    laws_block = (
        "SPĒKĀ ESOŠIE LIKUMI (OFICIĀLAIS TEKSTS):\n"
        "Analīzē balsties uz šiem likumu tekstiem. Atsaucies uz pantiem tikai tad, ja tie ir "
        "šajā tekstā. Ja neesi pārliecināts par panta numuru, nenorādi to un neizdomā.\n\n"
        + laws_text
    )
    try:
        cache = create_prompt_cache(api_key, localize_prompt(SYSTEM_PROMPT), laws_block)
        model = genai.GenerativeModel.from_cached_content(
            cached_content=cache, generation_config=generation_config
        )
        return model, source
    except Exception:
        # Nav atbalsta, par maz tokenu vai cita kļūda: darbs turpinās bez kešatmiņas
        return None, source


def generate_analysis(cached_model, model, active_prompt, parts):
    """Ģenerē analīzi ar kešatmiņu; ja tā nav pieejama, izmanto parasto pieprasījumu."""
    if cached_model is not None:
        try:
            return cached_model.generate_content(parts, stream=True)
        except Exception:
            pass
    return model.generate_content([active_prompt] + parts, stream=True)
@st.cache_data(ttl=datetime.timedelta(days=7), max_entries=100, show_spinner=False)
def analyze_contract(_api_key, lang_name, kind, content, nonce=0):
    """Veic pilnu analīzi (galvenā analīze + korektūra) un saglabā rezultātu atmiņā.
    Kešatmiņas atslēga = valoda + līguma saturs (teksts vai faila baiti) + nonce.
    Tas pats saturs = tūlītējs rezultāts bez API izsaukuma."""
    genai.configure(api_key=_api_key)
    cfg = genai.GenerationConfig(temperature=0, top_p=1, top_k=1)
    model = genai.GenerativeModel('gemini-3.8-flash', generation_config=cfg)
    cached_model, laws_source = get_cached_model(_api_key, cfg)
    active_prompt, _ = build_prompt()

    if kind == 'txt':
        parts = [f"LĪGUMA TEKSTS ANALĪZEI:\n{content}"]
        proof_input = f"{build_proof_prompt()}\n\nLĪGUMA TEKSTS:\n{content}"
    else:
        mime = {
            'pdf': 'application/pdf', 'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg'
        }.get(kind, 'application/octet-stream')
        file_part = {'mime_type': mime, 'data': content}
        parts = [file_part]
        proof_input = [build_proof_prompt(), file_part]

    def _collect(resp):
        return "".join(c.text for c in resp if hasattr(c, 'text') and c.text)

    main_text = linkify(_collect(generate_analysis(cached_model, model, active_prompt, parts)))
    if not main_text.strip():
        raise ValueError("Gemini atgrieza tukšu atbildi")  # kļūda netiek saglabāta kešatmiņā
    proof_text = _collect(model.generate_content(proof_input, stream=True))
    analysis_all = main_text + ("\n\n" + proof_text if proof_text.strip() else "")
    return analysis_all, laws_source + (" | ⚡ cache" if cached_model else "")

# ---------------------------------------------------------------
# Atsauces uz likumi.lv: [[DTIL 28]] -> klikšķināma saite uz attiecīgo pantu
# ---------------------------------------------------------------
LAW_LINKS = {
    "DTIL": "https://likumi.lv/ta/id/322216",
    "CL": "https://likumi.lv/ta/id/90220",
}
CITATION_RE = re.compile(r"\[\[\s*(DTIL|CL)\s+(\d{1,4})\s*\]\]")


def sources_md():
    return (
        f"**{T['sources_label']}** "
        f"[{T['law1_name']}]({LAW_LINKS['DTIL']}) · "
        f"[{T['law2_name']}]({LAW_LINKS['CL']})"
    )


def _citation_to_link(match):
    code, number = match.group(1), match.group(2)
    label = T["cite_dtil" if code == "DTIL" else "cite_cl"].format(n=number)
    return f"[{label}]({LAW_LINKS[code]}#p{number})"


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
    title = doc.add_heading(T["doc_title"], level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_meta = p_meta.add_run(
        f"{T['doc_prepared']}: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}"
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
            elif "Apsver" in line_str or "🔵" in line_str:
                color = RGBColor(37, 99, 235)
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
        self.cell(0, 8, T["pdf_header"], align="L")
        self.set_y(22)

    def footer(self):
        self.set_y(-15)
        self.set_font("DejaVu", "", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"{T['pdf_page']} {self.page_no()} | {T['pdf_footer']}", align="C")


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
    pdf.cell(0, 10, T["pdf_title"],
             new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"{T['pdf_date']}: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}",
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
            elif "Apsver" in raw or "🔵" in raw:
                pdf.set_text_color(37, 99, 235)
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
force_new = st.checkbox("♻️ Jauna analīze (neizmantot kešu) / New analysis (skip cache)", key="force_new")
analyze_button = st.button(T["analyze_btn"], type="primary")


UNSUPPORTED_PDF_SCRIPTS = re.compile(
    "[\u0590-\u08FF\u0900-\u0DFF\u0E00-\u0EFF\u1000-\u109F\u1780-\u17FF\u3000-\u9FFF\uAC00-\uD7AF]"
)


def build_downloads(report_text):
    """Izveido Word un PDF failus (baitos), lai tos varētu lejupielādēt vairākkārt."""
    result = {"docx": None, "pdf": None, "docx_err": None, "pdf_err": None}
    try:
        result["docx"] = generate_pretty_docx(report_text).getvalue()
    except Exception as e:
        result["docx_err"] = str(e)
    if UNSUPPORTED_PDF_SCRIPTS.search(report_text):
        # Fonts DejaVu neatbalsta ķīniešu, japāņu, korejiešu, arābu, hindi, taju u.c. rakstību
        result["pdf_err"] = "PDF is not available for this script (font limitation). Use the Word file. / Izmantojiet Word failu."
    else:
        try:
            result["pdf"] = generate_pretty_pdf(report_text).getvalue()
        except Exception as e:
            result["pdf_err"] = str(e)
    return result

# ---------------------------------------------------------------
# Kopsavilkuma panelis (kartītes) un sakļaujamās sadaļas
# ---------------------------------------------------------------
SECTION_CODES = {"🔴": "K", "🟡": "B", "🔵": "A", "⚪": "N"}


def split_sections(text):
    """Sadala analīzi sadaļās pēc virsrakstiem '## ...'."""
    sections, head, body = [], None, []
    for line in text.split("\n"):
        if line.startswith("## "):
            sections.append((head, "\n".join(body)))
            head, body = line[3:].strip(), []
        else:
            body.append(line)
    sections.append((head, "\n".join(body)))
    return [(h, b) for h, b in sections if h or b.strip()]


def count_items(head, body):
    """Saskaita punktus (K1, B1, A1, N1...) sadaļā pēc virsraksta emocijzīmes."""
    for emoji, letter in SECTION_CODES.items():
        if head and emoji in head:
            return len(re.findall(r"(?m)^\s*(?:[-*•]\s*)?[*_]*" + letter + r"\d+\b", body))
    return 0


def render_dashboard(text):
    totals = {"🔴": 0, "🟡": 0, "🔵": 0}
    for head, body in split_sections(text):
        for emoji in totals:
            if head and emoji in head:
                totals[emoji] += count_items(head, body)
    cards = [("🔴", "#dc2626", T["h_crit"]), ("🟡", "#d97706", T["h_warn"]), ("🔵", "#2563eb", T["h_border"])]
    for col, (emoji, color, label) in zip(st.columns(3), cards):
        col.markdown(
            f'<div style="border-left:6px solid {color};background:{color}22;border-radius:10px;padding:12px 14px;">'
            f'<div style="font-size:2rem;font-weight:700;color:{color};line-height:1.1">{totals[emoji]}</div>'
            f'<div style="font-size:0.85rem;opacity:0.9">{label}</div></div>',
            unsafe_allow_html=True,
        )


def render_sections(text):
    """Katra sadaļa ir sakļaujama. Pēc noklusējuma atvērtas: kopējais novērtējums, kritiskie, brīdinājumi."""
    for head, body in split_sections(text):
        if head is None:
            st.markdown(body)
            continue
        n = count_items(head, body)
        label = head + (f" ({n})" if n else "")
        with st.expander(label, expanded=any(e in head for e in ("📊", "🔴", "🟡"))):
            st.markdown(body)
# ---------------------------------------------------------------
# Pantu skaidrojumi: tooltip uz atsaucēm + pieskaramas kartes (bez Gemini izsaukumiem)
# Dati: (1) iebūvēta vārdnīca ar īsiem skaidrojumiem, (2) panta teksts no likumi.lv,
# kas jau tiek lejupielādēts analīzei (kešatmiņā)
# ---------------------------------------------------------------
ARTICLE_NOTES = {
    ("DTIL", 1): "Likuma mērķis ir mājokļa pieejamība, nodrošinot taisnīgu līdzsvaru starp izīrētāja un īrnieka interesēm.",
    ("DTIL", 2): "Likums regulē izīrētāja un īrnieka attiecības. Ko tas neregulē, to regulē Civillikums. Likums neattiecas, piemēram, uz īslaicīgu (sezonālu) lietošanu atpūtai un tūrismam.",
    ("DTIL", 3): "Dzīvojamās telpas īre ir telpas lietošanas tiesību nodošana citai personai par maksu.",
    ("DTIL", 5): "Īrnieks var būt tikai fiziska persona. Izīrētājs parasti ir telpas īpašnieks vai valdītājs. Vienu līgumu var slēgt ar vairākiem īrniekiem.",
    ("DTIL", 6): "Vienīgais telpas lietošanas pamats ir rakstveidā noslēgts īres līgums. Mutiska vienošanās nav pietiekama.",
    ("DTIL", 7): "Īres līgums jāslēdz rakstveidā.",
    ("DTIL", 8): "Līgumā obligāti jābūt: telpas adresei, platībai un raksturojumam; līguma termiņam; īres maksai, tās maksāšanas kārtībai un termiņiem, maksas grozīšanas nosacījumiem; saistītajiem pakalpojumiem un to apmaksai; ziņām par iemitinātajām personām. Ja kas no tā trūkst, līgums ir nepilnīgs un var radīt strīdus.",
    ("DTIL", 9): "Īres līgumu slēdz uz noteiktu termiņu.",
    ("DTIL", 10): "Īres maksu nosaka rakstveida vienošanās. Līguma laikā to drīkst mainīt tikai tad, ja līgumā ir paredzēti maksas grozīšanas nosacījumi un kārtība.",
    ("DTIL", 11): "Komunālo un citu ar telpas lietošanu saistīto maksājumu veikšanas kārtību nosaka īres līgums, tāpēc tā jāraksta skaidri.",
    ("DTIL", 12): "Drošības nauda nedrīkst pārsniegt divu mēnešu īres maksu. Īrnieka parādus var dzēst no tās (par to izīrētājam jāpaziņo nekavējoties). Neizmantoto daļu jāatdod ne vēlāk kā dienā, kad telpa tiek atbrīvota, ja līgumā nav noteikts citādi.",
    ("DTIL", 13): "Īrniekam ir tiesības netraucēti lietot telpu atbilstoši līgumam. Tāpēc izīrētāja apmeklējumu kārtība līgumā jānosaka skaidri.",
    ("DTIL", 14): "Īrnieks drīkst iemitināt laulāto un bērnus, iepriekš rakstveidā informējot izīrētāju. Citas personas tikai ar izīrētāja rakstveida piekrišanu. Ziņas par visām iemitinātajām personām jānorāda līgumā.",
    ("DTIL", 15): "Izīrētājs nedrīkst liegt īrniekam deklarēt īrēto telpu kā dzīvesvietu (arī iemitinātajām personām).",
    ("DTIL", 16): "Iemitinātajām personām nav patstāvīgu tiesību uz telpu. Īrnieka nāves gadījumā ģimenes loceklis 3 mēnešu laikā var prasīt īres līguma slēgšanu uz tiem pašiem nosacījumiem, bet jāsedz iepriekšējā īrnieka parādi.",
    ("DTIL", 17): "Īrnieks ievēro līgumu, maksā īri un ar telpu saistītos maksājumus, saudzīgi lieto telpu, neveic neatļautu pārbūvi vai remontu, nekavējoties ziņo par bojājumiem un līgumam beidzoties atbrīvo telpu. Vairāki īrnieki atbild solidāri.",
    ("DTIL", 18): "Izīrētājs nodod telpu un nodrošina netraucētu lietošanu visu līguma laiku, nekavējoties informē par savu adresi, kontaktiem un telpas atsavināšanu, un (ciktāl tas no viņa atkarīgs) uztur māju.",
    ("DTIL", 19): "Īres līgums izbeidzas: termiņam beidzoties, telpai gājot bojā, ar īrnieka nāvi vai kad īrnieks telpu iegūst īpašumā.",
    ("DTIL", 20): "Īrniekam ir tiesības vienpusēji atkāpties no īres līguma (brīdināšanas termiņš ir 27. pantā).",
    ("DTIL", 21): "Izīrētājs var izbeigt līgumu vai atkāpties no tā tikai likumā noteiktajos gadījumos (22.–26. pants).",
    ("DTIL", 22): "Izīrētājs var izbeigt līgumu, ja īrnieks vai iemitinātās personas bojā telpu vai māju, izmanto to līgumā neparedzētiem mērķiem vai pārkāpj lietošanas noteikumus.",
    ("DTIL", 23): "Izīrētājs var izbeigt līgumu, ja bez viņa piekrišanas telpu lieto līgumā neparedzētas personas.",
    ("DTIL", 24): "Izīrētājs var izbeigt līgumu, ja kavētā īre pārsniedz divu mēnešu īres maksu (vai līgumā noteikto lielāko apmēru) vai ar telpu saistīto maksājumu kavējums pārsniedz divus mēnešus.",
    ("DTIL", 25): "Izbeigšana mājas nojaukšanas vai pārbūves dēļ: izīrētājam jāatlīdzina pārcelšanās izdevumi, ieguldījumi un zaudējumi.",
    ("DTIL", 26): "Atkāpšanās no līguma, ja pieņemts lēmums aizliegt mājas ekspluatāciju: izīrētājam jāatlīdzina pārcelšanās izdevumi, ieguldījumi un zaudējumi.",
    ("DTIL", 27): "Brīdinājums par izbeigšanu jāsniedz rakstveidā. Īrnieks brīdina vismaz 1 mēnesi iepriekš (līgumiem līdz 10 gadiem). Izīrētāja termiņi atkarīgi no izbeigšanas pamata un līguma ilguma. Brīdinājumā jānorāda neizpildītā saistība, kavējuma periods, summa un sekas.",
    ("DTIL", 28): "Ja telpu pārdod, jaunajam īpašniekam līgums ir saistošs TIKAI tad, ja īres līgums ir ierakstīts zemesgrāmatā. Ja nav, jaunais īpašnieks var prasīt telpas atbrīvot 2 mēnešu laikā.",
    ("DTIL", 29): "Līgumam izbeidzoties, īrniekam un iemitinātajiem telpa jāatbrīvo. Ja neatbrīvo, jāatlīdzina zaudējumi un jāmaksā kompensācija par lietošanu.",
    ("DTIL", 35): "Īres tiesības var nostiprināt zemesgrāmatā. Tas aizsargā īrnieku, ja īpašnieks mainās (28. pants).",
    ("DTIL", 36): "Īres tiesību nostiprināšanai piemēro Zemesgrāmatu likumu, ciktāl šis likums nenosaka citādi.",
    ("DTIL", 37): "Īres tiesību nostiprināšana un grozīšana zemesgrāmatā ir atbrīvota no zemesgrāmatas kancelejas nodevas.",
    ("DTIL", 39): "Nostiprināšanu var lūgt izīrētājs un īrnieks kopā, vai viens no viņiem, ja otrs to ir pilnvarojis.",
    ("DTIL", 42): "Izīrētājam piemēro naudas sodu par traucējumiem, kas liedz īrniekam lietot telpu vai izmantot apkuri, elektrību, ūdeni, kanalizāciju vai atkritumu izvešanu.",
}

LINK_RE = re.compile(r"\[([^\]]+)\]\((https://likumi\.lv/ta/id/(322216|90220)#p(\d+))\)")


@st.cache_data(ttl=datetime.timedelta(hours=1), show_spinner=False)
def article_index(code):
    """Atgriež {panta numurs: teksts} no likumi.lv teksta (bez Gemini). Kļūdas gadījumā atgriež {}."""
    try:
        name, url, ranges = LAW_SOURCES[0] if code == "DTIL" else LAW_SOURCES[1]
        text = fetch_law(url)
    except Exception:
        return {}
    index, cur, buf = {}, None, []

    def flush():
        if cur is not None and buf:
            body = " ".join(buf).strip()
            if len(body) > len(index.get(cur, "")):
                index[cur] = body

    for line in text.split("\n"):
        if code == "DTIL":
            m = re.match(r"^(\d{1,2})\. pants\b", line)
            num = int(m.group(1)) if m else None
        else:
            m = ARTICLE_START.match(line)
            num = int(m.group(1) or m.group(2)) if m else None
            if num is not None and ranges and not any(lo <= num <= hi for lo, hi in ranges):
                flush()
                cur, buf = None, []
                continue
        if num is not None:
            flush()
            cur, buf = num, [line]
        elif cur is not None:
            buf.append(line)
    flush()
    return index


def article_excerpt(code, num, limit=600):
    """Oficiālā panta teksta sākums (no likumi.lv)."""
    raw = article_index(code).get(num, "")
    raw = re.sub(r"^\d+\.?\s*(?:pants\.?)?\s*", "", raw)
    raw = re.sub(r"\s+", " ", raw).strip()
    return raw if (limit is None or len(raw) <= limit) else raw[:limit].rsplit(" ", 1)[0] + "…"


def article_tooltip(code, num):
    tip = ARTICLE_NOTES.get((code, num)) or article_excerpt(code, num, 220)
    tip = tip.replace('"', "'").replace("\\", "").replace("\n", " ")
    return tip if len(tip) <= 300 else tip[:300].rsplit(" ", 1)[0] + "…"


def add_tooltips(text):
    """Pievieno atsauču saitēm paskaidrojumu (parādās, uzvedot peli virs saites)."""
    def _sub(m):
        code = "DTIL" if m.group(3) == "322216" else "CL"
        tip = article_tooltip(code, int(m.group(4)))
        return f'[{m.group(1)}]({m.group(2)} "{tip}")' if tip else m.group(0)
    return LINK_RE.sub(_sub, text)


def render_law_help(text):
    """Pieskaramas kartes ar skaidrojumu katram analīzē minētajam pantam (der arī telefonam)."""
    seen = []
    for m in LINK_RE.finditer(text):
        key = ("DTIL" if m.group(3) == "322216" else "CL", int(m.group(4)))
        if key not in seen:
            seen.append(key)
    if not seen:
        return
    st.markdown("### 📖 Pantu skaidrojumi (pieskarieties pantam)")
    cols = st.columns(2)
    for i, (code, num) in enumerate(seen):
        label = T["cite_dtil" if code == "DTIL" else "cite_cl"].format(n=num)
        with cols[i % 2]:
            box = st.popover(label, use_container_width=True) if hasattr(st, "popover") else st.expander(label)
            with box:
                note = ARTICLE_NOTES.get((code, num))
                excerpt = article_excerpt(code, num, None)
                if note:
                    st.markdown("**Ko tas nozīmē īrniekam:** " + note)
                if excerpt:
                    st.markdown("**Oficiālais teksts:** " + excerpt)
                if not note and not excerpt:
                    st.markdown("Skaidrojums nav pieejams. Atveriet pilno tekstu likumi.lv.")
                st.markdown(f"[Atvērt likumi.lv]({LAW_LINKS[code]}#p{num})")
                st.caption("Informatīvs skaidrojums pēc likuma teksta, nav juridiska konsultācija.")
# ---------------------------------------------------------------
# Paskaidrojuma lodziņš uz pantu atsaucēm: peles uzvešana (dators) vai ilga turēšana (telefons/planšete)
# Parasts pieskāriens joprojām atver pantu likumi.lv. Strādā ar JavaScript pārlūkā (bez Gemini).
# ---------------------------------------------------------------
LAW_TIP_JS = """
<script>
(function () {
  var P = window.parent, D = P.document;
  P.__lawTips = __MAP__;                  // skaidrojumi tiek atjaunoti katru reizi
  if (P.__lawTipsReady) return;           // apstrādātāji tiek pievienoti tikai vienu reizi
  P.__lawTipsReady = true;

  var st = D.createElement("style");
  st.textContent =
    "a[href*='likumi.lv/ta/id/'] { -webkit-touch-callout: none; -webkit-user-select: none; user-select: none; }" +
    "#law-tip-bubble { position: fixed; display: none; z-index: 2147483647; max-width: min(92vw, 480px); max-height: 60vh;" +
    " overflow: auto; background: #1f2430; color: #f3f4f6; border: 1px solid #4b5160; border-radius: 12px;" +
    " padding: 12px 14px; font: 14px/1.5 sans-serif; white-space: pre-wrap; box-shadow: 0 8px 30px rgba(0,0,0,0.45); }";
  D.head.appendChild(st);
  var bubble = D.createElement("div");
  bubble.id = "law-tip-bubble";
  D.body.appendChild(bubble);

  function infoOf(el) {
    var a = el && el.closest ? el.closest("a") : null;
    if (!a) return null;
    var m = (a.getAttribute("href") || "").match(/likumi\\.lv\\/ta\\/id\\/(\\d+)#p(\\d+)/);
    return m ? { a: a, k: m[1] + "#p" + m[2] } : null;
  }
  function show(info) {
    var t = P.__lawTips[info.k];
    if (!t) return false;
    info.a.removeAttribute("title");
    bubble.textContent = t;
    bubble.style.display = "block";
    var r = info.a.getBoundingClientRect(), w = bubble.offsetWidth, h = bubble.offsetHeight;
    var left = Math.max(8, Math.min(r.left, P.innerWidth - w - 8));
    var top = r.bottom + 8;
    if (top + h > P.innerHeight - 8) top = Math.max(8, r.top - h - 8);
    bubble.style.left = left + "px";
    bubble.style.top = top + "px";
    return true;
  }
  function hide() { bubble.style.display = "none"; }

  var timer = null, held = false, lastTouch = 0, sx = 0, sy = 0;
  D.addEventListener("touchstart", function (e) {
    lastTouch = Date.now();
    if (bubble.contains(e.target)) return;
    clearTimeout(timer);
    held = false;
    var info = infoOf(e.target);
    hide();
    if (!info) return;
    var t = e.touches[0];
    sx = t.clientX; sy = t.clientY;
    timer = setTimeout(function () {
      held = show(info);
      if (held) setTimeout(function () { held = false; }, 800);
    }, 450);
  }, { passive: true });
  D.addEventListener("touchmove", function (e) {
    var t = e.touches[0];
    if (Math.abs(t.clientX - sx) > 10 || Math.abs(t.clientY - sy) > 10) clearTimeout(timer);
  }, { passive: true });
  D.addEventListener("touchend", function () { clearTimeout(timer); }, { passive: true });
  D.addEventListener("touchcancel", function () { clearTimeout(timer); }, { passive: true });
  D.addEventListener("contextmenu", function (e) { if (infoOf(e.target)) e.preventDefault(); });
  D.addEventListener("click", function (e) {
    if (held && infoOf(e.target)) { e.preventDefault(); e.stopPropagation(); held = false; return; }
    if (!bubble.contains(e.target)) hide();
  }, true);
  D.addEventListener("mouseover", function (e) {
    if (Date.now() - lastTouch < 1000) return;
    var info = infoOf(e.target);
    if (info) show(info);
  });
  D.addEventListener("mouseout", function (e) {
    if (Date.now() - lastTouch < 1000) return;
    if (infoOf(e.target)) hide();
  });
})();
</script>
"""


def law_tip_map(text):
    """Saliek pilnu skaidrojuma tekstu (bez saīsinājumiem) katram analīzē minētajam pantam."""
    tips = {}
    for m in LINK_RE.finditer(text):
        code = "DTIL" if m.group(3) == "322216" else "CL"
        num = int(m.group(4))
        key = f"{m.group(3)}#p{num}"
        if key in tips:
            continue
        parts = [T["cite_dtil" if code == "DTIL" else "cite_cl"].format(n=num)]
        note = ARTICLE_NOTES.get((code, num))
        full = article_excerpt(code, num, None)
        if note:
            parts.append("Ko tas nozīmē īrniekam: " + note)
        if full:
            parts.append("Oficiālais teksts: " + full)
        if not note and not full:
            parts.append("Skaidrojums nav pieejams. Atveriet pilno tekstu likumi.lv.")
        tips[key] = "\n\n".join(parts)
    return tips


def inject_law_tips(text):
    payload = json.dumps(law_tip_map(text), ensure_ascii=False).replace("</", "<\\/")
    components.html(LAW_TIP_JS.replace("__MAP__", payload), height=0)
# ---------------------------------------------------------------
# Soda naudas un procentu kalkulators (bez Gemini izsaukumiem)
# Likumi: Civillikums 1716. (līgumsoda griesti 10%), 1717. (samērīgums), 1722. (līgumsods un procenti),
# 1724.1 (tiesa samazina pārmērīgu līgumsodu), 1764. un 1765. (procenti, likumiskie procenti 6% gadā)
# ---------------------------------------------------------------
RATE_DAY_RE = re.compile(
    r"(\d+(?:[.,]\d+)?)\s*(?:%|procent\w*)[^.;\n%]{0,160}?(?:par\s+katru\s+(?:\w+\s+){0,3}?dienu|dienā)",
    re.IGNORECASE,
)
RENT_NUM_FIRST_RE = re.compile(r"īres\s+maks\w*[^\n]{0,120}?(\d[\d ]*(?:[.,]\d{1,2})?)\s*(?:EUR|eiro|€)", re.IGNORECASE)
RENT_CUR_FIRST_RE = re.compile(r"īres\s+maks\w*[^\n]{0,120}?(?:EUR|€)\s*(\d[\d ]*(?:[.,]\d{1,2})?)", re.IGNORECASE)
CAP_RE = re.compile(
    r"(?:ne\s+vairāk\s+kā|ne\s+vairāk\s+par|nepārsniedz\w*|līdz)\s+(?:\w+\s+){0,3}?10\s*(?:\(desmit\)\s*)?(?:%|procent)",
    re.IGNORECASE,
)


def _to_float(raw):
    try:
        return float(raw.replace(" ", "").replace(",", "."))
    except Exception:
        return None


def extract_terms(text):
    """Mēģina no līguma teksta nolasīt līgumsoda likmi dienā, īres maksu un 10% griestus (bez Gemini)."""
    terms = {}
    m = RATE_DAY_RE.search(text)
    if m:
        rate = _to_float(m.group(1))
        if rate and 0 < rate <= 100:
            terms["rate_day"] = rate
    for rx in (RENT_NUM_FIRST_RE, RENT_CUR_FIRST_RE):
        m = rx.search(text)
        if m:
            rent = _to_float(m.group(1))
            if rent and rent > 0:
                terms["rent"] = rent
                break
    if CAP_RE.search(text):
        terms["cap"] = True
    return terms


def eur(x):
    return (f"{x:,.2f}".replace(",", " ").replace(".", ",")) + " EUR"


def pct(x):
    return f"{x:g}".replace(".", ",")


def _cl(n, text=None):
    return f"[{text or f'Civillikuma {n}. pants'}]({LAW_LINKS['CL']}#p{n})"


def penalty_calc(principal, days, rate, daily, cap_declared):
    """Aprēķins un vērtējums. daily=True: likme % dienā (līgumsods); False: likme % gadā (nokavējuma procenti)."""
    statutory = principal * 0.06 * days / 365          # likumiskie procenti 6% gadā (CL 1765.)
    cap = principal * 0.10                             # līgumsoda griesti (CL 1716. 3. daļa)
    if daily:
        total = principal * rate / 100 * days
        annual = rate * 365
    else:
        total = principal * rate / 100 * days / 365
        annual = rate
    multiple = annual / 6
    res = {"total": total, "cap": cap, "statutory": statutory, "annual": annual, "multiple": multiple,
           "daily": daily, "excess": 0.0, "lawful": total, "days_to_cap": None, "claimable": None}

    if daily:
        rate_day = rate
        res["days_to_cap"] = 10 / rate_day
        res["lawful"] = min(total, cap)
        res["excess"] = max(0.0, total - cap)
        res["claimable"] = max(0.0, res["lawful"] - statutory)      # CL 1722.
        if total > cap + 0.005:
            res["level"] = "error"
            res["title"] = "🔴 Pārsniedz likumā noteikto"
            res["messages"] = [
                f"Pēc līguma aprēķinātais līgumsods ({eur(total)}) pārsniedz 10% no parāda ({eur(cap)}). "
                f"Saskaņā ar {_cl(1716, 'Civillikuma 1716. panta')} trešo daļu līgumsods par saistības nepienācīgu izpildi (kavējumu) "
                f"kopumā nevar pārsniegt 10% no pamatparāda. Pārsnieguma daļa ({eur(res['excess'])}) nav prasāma."
            ]
            return res
        msgs = []
        if not cap_declared:
            msgs.append(
                f"Līgumā nav noteikts kopējais līgumsoda griesti. Ar likmi {pct(rate)}% dienā 10% robeža tiks sasniegta pēc "
                f"{res['days_to_cap']:.0f} dienām. Saskaņā ar {_cl(1716, 'Civillikuma 1716. panta')} līgumā jābūt norādītam, ka līgumsods kopumā nepārsniedz 10%."
            )
        if rate > 0.1:
            msgs.append(
                f"Likme {pct(rate)}% dienā atbilst ~{annual:.0f}% gadā, t.i. {multiple:.1f} reižu vairāk nekā likumiskie procenti "
                f"(6% gadā, {_cl(1765)}). Likums neparedz konkrētu dienas likmes robežu, bet tiesa samazina pārmērīgu līgumsodu "
                f"({_cl(1717)}, {_cl(1724, 'Civillikuma 1724.¹ pants')}). 0,1% dienā ir biežāk sastopams samērīgs līmenis (informatīvs orientieris, nav likuma norma)."
            )
        if msgs:
            res["level"] = "warning"
            res["title"] = "🟡 Uzmanību"
            res["messages"] = msgs
        else:
            res["level"] = "success"
            res["title"] = "🟢 Atbilst normām"
            res["messages"] = [
                f"Līgumsods ({eur(total)}) nepārsniedz 10% griestus ({eur(cap)}) saskaņā ar {_cl(1716, 'Civillikuma 1716. panta')}, līgumā griesti ir noteikti, "
                f"un likme nepārsniedz 0,1% dienā."
            ]
        return res

    # nokavējuma procenti (% gadā)
    if rate <= 6:
        res["level"], res["title"] = "success", "🟢 Atbilst normām"
        res["messages"] = [f"Likme ({pct(rate)}% gadā) nepārsniedz likumiskos procentus (6% gadā, {_cl(1765)})."]
    elif rate <= 36.5:
        res["level"], res["title"] = "warning", "🟡 Uzmanību"
        res["messages"] = [
            f"Likme ({pct(rate)}% gadā) ir {multiple:.1f} reižu augstāka par likumiskajiem procentiem (6% gadā, {_cl(1765)}). "
            f"Civillikums neparedz konkrētu augšējo robežu, bet nesamērīgi un godīgai darījumu praksei neatbilstoši procenti "
            f"ir prettiesiski ({_cl(1764)}, {_cl(1717)})."
        ]
    else:
        res["level"], res["title"] = "error", "🔴 Risks"
        res["messages"] = [
            f"Likme ({pct(rate)}% gadā, ~{pct(round(rate / 365, 2))}% dienā) ir ļoti augsta ({multiple:.1f} reizes augstāka par likumiskajiem "
            f"procentiem) un var tikt atzīta par nesamērīgu un prettiesisku ({_cl(1764)}, {_cl(1717)})."
        ]
    return res


def render_penalty_calculator():
    """Interaktīvs kalkulators. Lauki tiek aizpildīti no līguma teksta, ja tas tika atrasts."""
    terms = st.session_state.get("terms") or {}
    sig = f"{terms.get('rent')}_{terms.get('rate_day')}_{terms.get('cap')}"
    with st.expander("🧮 Soda naudas un procentu kalkulators", expanded=False):
        if terms:
            st.caption("Lauki aizpildīti automātiski no līguma teksta (ja tas tika atrasts). Pārbaudiet un, ja vajag, izlabojiet.")
        kind = st.radio(
            "Sankcijas veids",
            ["Līgumsods / nokavējuma nauda (% dienā)", "Nokavējuma procenti (% gadā)"],
            key=f"calc_kind_{sig}",
        )
        daily = kind.startswith("Līgumsods")
        c1, c2, c3 = st.columns(3)
        principal = c1.number_input(
            "Parāda summa / īres maksa (EUR)", min_value=0.0, value=float(terms.get("rent") or 500.0),
            step=10.0, format="%.2f", key=f"calc_p_{sig}",
        )
        days = c2.number_input("Kavēto dienu skaits", min_value=0, value=30, step=1, key=f"calc_d_{sig}")
        rate = c3.number_input(
            "Likme līgumā (% dienā)" if daily else "Likme līgumā (% gadā)", min_value=0.0,
            value=float(terms.get("rate_day") or 0.1) if daily else 6.0,
            step=0.01, format="%.3f", key=f"calc_r_{sig}_{int(daily)}",
        )
        cap_declared = True
        if daily:
            cap_declared = st.checkbox(
                "Līgumā ir noteikts kopējais līgumsoda griesti (ne vairāk kā 10% no parāda)",
                value=bool(terms.get("cap")), key=f"calc_cap_{sig}",
            )
        if principal <= 0 or days <= 0 or rate <= 0:
            st.info("Ievadiet summu, dienu skaitu un likmi, lai redzētu aprēķinu.")
            return

        r = penalty_calc(principal, days, rate, daily, cap_declared)
        m1, m2, m3 = st.columns(3)
        m1.metric("Pēc līguma noteikumiem", eur(r["total"]))
        if daily:
            m2.metric("Likumā pieļaujamais maksimums (10%)", eur(r["cap"]))
        else:
            m2.metric("Starpība pret likumiskajiem", eur(max(0.0, r["total"] - r["statutory"])))
        m3.metric("Likumiskie procenti (6% gadā)", eur(r["statutory"]))

        body = "\n\n".join(f"- {m}" for m in r["messages"])
        {"success": st.success, "warning": st.warning, "error": st.error}[r["level"]](f"**{r['title']}**\n\n{body}")

        if daily:
            st.markdown(
                f"**Ja izīrētājs prasa gan likumiskos procentus, gan līgumsodu** (saskaņā ar {_cl(1722, 'Civillikuma 1722. pantu')}): līgumsodu var prasīt tikai "
                f"tādā apmērā, kādā tas pārsniedz prasīto procentu summu ({eur(r['statutory'])}). "
                f"Līgumsods šajā gadījumā: ne vairāk kā **{eur(r['claimable'])}**."
            )
        st.caption(
            "Aprēķins: vienkāršie procenti (bez procentu procentiem). Informatīvs aprēķins, nav juridiska konsultācija. "
            "Ja izīrētājs ir uzņēmējs vai līgums ir patērētāju līgums, var būt papildu ierobežojumi."
        )
def render_downloads(result):
    st.success(T["success"])
    st.markdown("---")
    st.markdown(T["dl_header"])
    col1, col2 = st.columns(2)
    today = datetime.date.today()

    with col1:
        if result["docx"]:
            st.download_button(
                label=T["dl_word"],
                data=result["docx"],
                file_name=f"Ires_Liguma_Analize_{today}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="dl_docx"
            )
        else:
            st.warning(f"{T['word_fail']}: {result['docx_err']}")

    with col2:
        if result["pdf"]:
            st.download_button(
                label=T["dl_pdf"],
                data=result["pdf"],
                file_name=f"Ires_Liguma_Analize_{today}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="dl_pdf"
            )
        else:
            st.warning(f"{T['pdf_fail']}: {result['pdf_err']}")


# ---------------------------------------------------------------
# Analīzes izpilde
# ---------------------------------------------------------------
if analyze_button:
    # Jauna analīze: izdzēš iepriekšējo saglabāto rezultātu
    st.session_state.pop("result", None)

    if not api_key:
        st.error(T["err_no_key"])
    elif uploaded_file is None and not pasted_text.strip():
        st.error(T["err_no_file"])
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
            # Ja ielīmēts teksts, izmanto to; citādi - augšupielādēto failu
            if pasted_text.strip():
                file_extension = "txt"
            else:
                file_extension = uploaded_file.name.split('.')[-1].lower()
                            # Sagatavo saturu: teksts (docx/ielīmēts) vai faila baiti (pdf/attēls)
            if file_extension in ('docx', 'txt'):
                kind = 'txt'
                if file_extension == 'txt':
                    content = pasted_text.strip()
                else:
                    content = "\n".join(
                        [p.text for p in docx.Document(io.BytesIO(uploaded_file.getvalue())).paragraphs if p.text.strip()]
                    )
            else:
                kind, content = file_extension, uploaded_file.getvalue()
            st.session_state["terms"] = extract_terms(content) if kind == 'txt' else {}

            _t0 = datetime.datetime.now()
            with st.spinner(T["spinner_analysis"]):
                analysis_all, laws_source = analyze_contract(
                    api_key, T["report_lang_name"], kind, content,
                    datetime.datetime.now().timestamp() if force_new else 0,
                )
            if (datetime.datetime.now() - _t0).total_seconds() < 2:
                laws_source += " | 💾 0 cost (cache)"
            st.caption(f"{T['law_source']}: {laws_source}")
            st.markdown(analysis_all)

            if analysis_all:
                st.markdown(sources_md())
                report_text = analysis_all + "\n\n## " + T["sources_head"] + "\n" + sources_md()
                result = build_downloads(report_text)
                result.update({
                    "analysis": analysis_all,
                    "laws_source": laws_source,
                })
                # Saglabā, lai analīze nepazustu pēc lejupielādes
                st.session_state["result"] = result
                render_downloads(result)

        except Exception as e:
            st.error(f"{T['err_process']}: {str(e)}")
        if "result" in st.session_state:
            st.rerun()

elif "result" in st.session_state:
    # Lapa pārlādējās (piem., nospiesta lejupielādes poga): parāda saglabāto analīzi
    saved = st.session_state["result"]
    st.caption(f"{T['law_source']}: {saved['laws_source']}")
    render_dashboard(saved["analysis"])
    render_sections(saved["analysis"])
    render_law_help(saved["analysis"])
    inject_law_tips(saved["analysis"])
    st.markdown(sources_md())
    render_downloads(saved)

render_penalty_calculator()
