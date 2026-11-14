import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import io
import re
import html as _html
import json
import datetime
import os
import requests
from bs4 import BeautifulSoup
import analysis_engine as ae

# ---------------------------------------------------------------
# Lapas konfigurācija
# ---------------------------------------------------------------
st.set_page_config(
    page_title="Īres un nomas līgumu analizētājs",
    page_icon="🏠",
    layout="centered",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------
# Tīrs, komerciāls izskats: paslēpj Streamlit kājeni ("Made with Streamlit"),
# trīs svītru izvēlni (hamburger), "Deploy" pogu un krāsaino joslu augšā.
# Pati galvene paliek (caurspīdīga), jo tajā ir sānu paneļa poga.
# ---------------------------------------------------------------
st.markdown(
    """
<style>
#MainMenu, [data-testid="stMainMenu"], [data-testid="stToolbarActions"],
[data-testid="stAppDeployButton"], .stDeployButton,
[data-testid="stDecoration"], [data-testid="stStatusWidget"],
footer, [data-testid="stFooter"], .viewerBadge_container__1QSob,
[class^="viewerBadge_"], [class*=" viewerBadge_"] {
    display: none !important;
    visibility: hidden !important;
}
[data-testid="stHeader"] { background: transparent !important; box-shadow: none !important; }
[data-testid="stMainBlockContainer"], .stMainBlockContainer, .block-container {
    padding-top: 0.4rem !important;
}
[data-testid="stMainBlockContainer"] h1 { padding-top: 0 !important; margin-top: 0 !important; }
@media (max-width: 640px) {
    [data-testid="stMainBlockContainer"] h1 {
        font-size: 1.45rem !important; line-height: 1.2 !important;
        overflow-wrap: anywhere; word-break: normal;
    }
}
</style>
""",
    unsafe_allow_html=True,
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
        "live_provisional": "🔄 Analīze vēl notiek. Šobrīd redzamā analīze ir provizoriska: tā vēl tiek pārbaudīta un papildināta.",
        "live_done": "✅ Analīze ir beigusies.",
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
        "live_provisional": "🔄 The analysis is still running. What you see now is provisional: it is still being verified and extended.",
        "live_done": "✅ The analysis has finished.",
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
        "live_provisional": "🔄 Анализ ещё выполняется. Показанный сейчас результат предварительный: он ещё проверяется и дополняется.",
        "live_done": "✅ Анализ завершён.",
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
        "h_overall": "Kopējais novērtējums", "h_crit": "Riska faktori (finansiāli un citi zaudējumi)", "h_warn": "Brīdinājumi (likuma pārkāpumi, neierakstīti vai neskaidri punkti)",
        "h_recs": "Ieteiktie labojumi (kā labot katru riska faktoru un brīdinājumu)", "h_border": "Apsveramie punkti (ieteicams pievienot vai precizēt)", "h_neutral": "Neitrālās piezīmes (citi nosacījumi un gramatikas kļūdas)",
        "h_money": "Naudas nosacījumi (neitrāls saraksts bez vērtējuma)",
        "no_proof_errors": "Korektūras kļūdu netika atrasts.",
        "label_proposed": "Piedāvātais formulējums", "label_translation": "Tulkojums", "label_fix": "Labojums",
        "disclaimer": "Piezīme: šī analīze ir automātiska pārbaude pirms parakstīšanas, nevis juridiska konsultācija. Svarīgiem vai dārgiem līgumiem konsultējieties ar juristu.",
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
        "h_overall": "Overall assessment", "h_crit": "Risk factors (financial and other losses)", "h_warn": "Warnings (legal violations, missing or unclear points)",
        "h_recs": "Suggested corrections (how to fix each point found)", "h_border": "Points to consider (recommended to add or clarify)", "h_neutral": "Neutral notes (other terms and grammar errors)",
        "h_money": "Money terms (neutral list, no judgement)",
        "no_proof_errors": "No proofreading errors were found.",
        "label_proposed": "Proposed wording", "label_translation": "Translation", "label_fix": "Correction",
        "disclaimer": "Note: this analysis is an automated check before signing, not legal advice. For important or high-value contracts, consult a lawyer.",
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
        "h_overall": "Общая оценка", "h_crit": "Факторы риска (финансовые и другие убытки)", "h_warn": "Предупреждения (нарушения закона, отсутствующие или неясные пункты)",
        "h_recs": "Рекомендуемые исправления (как исправить каждый найденный пункт)", "h_border": "Пункты для рассмотрения (рекомендуется добавить или уточнить)", "h_neutral": "Нейтральные замечания (другие условия и грамматические ошибки)",
        "h_money": "Денежные условия (нейтральный список, без оценки)",
        "no_proof_errors": "Ошибок при корректуре не найдено.",
        "label_proposed": "Предлагаемая формулировка", "label_translation": "Перевод", "label_fix": "Исправление",
        "disclaimer": "Примечание: этот анализ — автоматическая проверка перед подписанием, а не юридическая консультация. Для важных или дорогостоящих договоров обратитесь к юристу.",
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

# ---------------------------------------------------------------
# Tulkojumi: pārējais saskarnes teksts (testa režīms, vēsture, pantu skaidrojumi, avoti)
# ---------------------------------------------------------------
TR_NEW2 = {
    "Latviešu": {
        "test_mode_label": "🧪 Testa režīms (bez Gemini, 0 €): iebūvēts paraugs",
        "force_new_label": "♻️ Jauna analīze (neizmantot kešu)",
        "hist_label": "📜 Sesijas vēsture (iepriekšējās analīzes šajā sesijā)",
        "hist_test": "Testa paraugs",
        "hist_pasted": "Ielīmēts teksts",
        "law_head": "📖 Pantu skaidrojumi (pieskarieties pantam)",
        "law_lv_note": "",
        "law_means": "Ko tas nozīmē īrniekam",
        "law_official": "Oficiālais teksts",
        "law_none": "Skaidrojums nav pieejams. Atveriet pilno tekstu likumi.lv.",
        "law_open": "Atvērt likumi.lv",
        "law_disclaimer": "Informatīvs skaidrojums pēc likuma teksta, nav juridiska konsultācija.",
        "test_laws_source": "🧪 TESTA REŽĪMS: iebūvēts paraugs, Gemini netika izsaukts (0.00$)",
        "test_note": "🧪 Testa režīms: iebūvēts paraugs, Gemini netika izsaukts (0.00$).",
        "src_auto": "likumi.lv (automātiski, {n} no {m} likumiem)",
        "src_backup": "rezerves fails likumi.txt",
        "src_none": "nav pieejams",
        "src_failed": "neizdevās",
    },
    "English": {
        "test_mode_label": "🧪 Test mode (no Gemini, €0): built-in sample",
        "force_new_label": "♻️ New analysis (skip cache)",
        "hist_label": "📜 Session history (earlier analyses in this session)",
        "hist_test": "Test sample",
        "hist_pasted": "Pasted text",
        "law_head": "📖 Article explanations (tap an article)",
        "law_lv_note": "The explanations and the official text are in Latvian.",
        "law_means": "What it means for the tenant",
        "law_official": "Official text",
        "law_none": "No explanation available. Open the full text on likumi.lv.",
        "law_open": "Open on likumi.lv",
        "law_disclaimer": "Informational explanation based on the text of the law, not legal advice.",
        "test_laws_source": "🧪 TEST MODE: built-in sample, Gemini was not called ($0.00)",
        "test_note": "🧪 Test mode: built-in sample, Gemini was not called ($0.00).",
        "src_auto": "likumi.lv (automatic, {n} of {m} laws)",
        "src_backup": "backup file likumi.txt",
        "src_none": "not available",
        "src_failed": "failed",
    },
    "Русский": {
        "test_mode_label": "🧪 Тестовый режим (без Gemini, 0 €): встроенный образец",
        "force_new_label": "♻️ Новый анализ (не использовать кэш)",
        "hist_label": "📜 История сессии (предыдущие анализы в этой сессии)",
        "hist_test": "Тестовый образец",
        "hist_pasted": "Вставленный текст",
        "law_head": "📖 Пояснения к статьям (нажмите на статью)",
        "law_lv_note": "Пояснения и официальный текст приведены на латышском языке.",
        "law_means": "Что это значит для арендатора",
        "law_official": "Официальный текст",
        "law_none": "Пояснение недоступно. Откройте полный текст на likumi.lv.",
        "law_open": "Открыть на likumi.lv",
        "law_disclaimer": "Информационное пояснение по тексту закона, не юридическая консультация.",
        "test_laws_source": "🧪 ТЕСТОВЫЙ РЕЖИМ: встроенный образец, Gemini не вызывался ($0.00)",
        "test_note": "🧪 Тестовый режим: встроенный образец, Gemini не вызывался ($0.00).",
        "src_auto": "likumi.lv (автоматически, {n} из {m} законов)",
        "src_backup": "резервный файл likumi.txt",
        "src_none": "недоступно",
        "src_failed": "не удалось",
    },
}
for _l, _extra in TR_NEW2.items():
    TR[_l].update(_extra)
# ---------------------------------------------------------------
# Tulkojumi: sesija, kalkulators un čats (latviešu, angļu, krievu). Citām valodām tos iztulko Gemini
# (jo TR["English"] ir pamats tulkošanai). Veidnēs {vārdi} ir vietturi, kas jāsaglabā.
# ---------------------------------------------------------------
TR_NEW = {
    "Latviešu": {
        "reset_btn": "🔄 Notīrīt sesiju / Sākt no jauna",
        "sidebar_toggle": "Rādīt / paslēpt iestatījumu paneli",
        "tab_risks": "📊 Risku analīze",
        "tab_calc": "🧮 Soda kalkulators",
        "tab_docs": "📄 Dokumentu lejupielāde",
        "tab_chat": "💬 Čats ar līgumu",
        "tab_fixes": "### 🟢 Ieteikumi un labojumi",
        "cite_cl_gen": "Civillikuma {n}. panta",
        "cite_cl_acc": "Civillikuma {n}. pantu",
        # --- kalkulators
        "calc_title": "🧮 Soda naudas un procentu kalkulators",
        "calc_autofill": "Lauki aizpildīti automātiski no līguma teksta (ja tas tika atrasts). Pārbaudiet un, ja vajag, izlabojiet.",
        "calc_kind": "Sankcijas veids",
        "calc_kind_daily": "Līgumsods / nokavējuma nauda (% dienā)",
        "calc_kind_annual": "Nokavējuma procenti (% gadā)",
        "calc_principal": "Parāda summa / īres maksa (EUR)",
        "calc_days": "Kavēto dienu skaits",
        "calc_rate_day": "Likme līgumā (% dienā)",
        "calc_rate_year": "Likme līgumā (% gadā)",
        "calc_cap": "Līgumā ir noteikts kopējais līgumsoda griesti (ne vairāk kā 10% no parāda)",
        "calc_enter": "Ievadiet summu, dienu skaitu un likmi, lai redzētu aprēķinu.",
        "calc_m_contract": "Pēc līguma noteikumiem",
        "calc_m_cap": "Likumā pieļaujamais maksimums (10%)",
        "calc_m_diff": "Starpība pret likumiskajiem",
        "calc_m_stat": "Likumiskie procenti (6% gadā)",
        "calc_t_error": "🔴 Pārsniedz likumā noteikto",
        "calc_t_warn": "🟡 Uzmanību",
        "calc_t_ok": "🟢 Atbilst normām",
        "calc_t_risk": "🔴 Risks",
        "calc_msg_exceed": "Pēc līguma aprēķinātais līgumsods ({total}) pārsniedz 10% no parāda ({cap}). Saskaņā ar {cl1716_gen} trešo daļu līgumsods par saistības nepienācīgu izpildi (kavējumu) kopumā nevar pārsniegt 10% no pamatparāda. Pārsnieguma daļa ({excess}) nav prasāma.",
        "calc_msg_nocap": "Līgumā nav noteikts kopējais līgumsoda griesti. Ar likmi {rate}% dienā 10% robeža tiks sasniegta pēc {days_to_cap} dienām. Saskaņā ar {cl1716_gen} līgumā jābūt norādītam, ka līgumsods kopumā nepārsniedz 10%.",
        "calc_msg_highrate": "Likme {rate}% dienā atbilst ~{annual}% gadā, t.i. {multiple} reižu vairāk nekā likumiskie procenti (6% gadā, {cl1765}). Likums neparedz konkrētu dienas likmes robežu, bet tiesa samazina pārmērīgu līgumsodu ({cl1717}, {cl1724}). 0,1% dienā ir biežāk sastopams samērīgs līmenis (informatīvs orientieris, nav likuma norma).",
        "calc_msg_ok_daily": "Līgumsods ({total}) nepārsniedz 10% griestus ({cap}) saskaņā ar {cl1716_gen}, līgumā griesti ir noteikti, un likme nepārsniedz 0,1% dienā.",
        "calc_msg_ok_annual": "Likme ({rate}% gadā) nepārsniedz likumiskos procentus (6% gadā, {cl1765}).",
        "calc_msg_warn_annual": "Likme ({rate}% gadā) ir {multiple} reižu augstāka par likumiskajiem procentiem (6% gadā, {cl1765}). Civillikums neparedz konkrētu augšējo robežu, bet nesamērīgi un godīgai darījumu praksei neatbilstoši procenti ir prettiesiski ({cl1764}, {cl1717}).",
        "calc_msg_risk_annual": "Likme ({rate}% gadā, ~{rate_day}% dienā) ir ļoti augsta ({multiple} reizes augstāka par likumiskajiem procentiem) un var tikt atzīta par nesamērīgu un prettiesisku ({cl1764}, {cl1717}).",
        "calc_both": "**Ja izīrētājs prasa gan likumiskos procentus, gan līgumsodu** (saskaņā ar {cl1722_acc}): līgumsodu var prasīt tikai tādā apmērā, kādā tas pārsniedz prasīto procentu summu ({statutory}). Līgumsods šajā gadījumā: ne vairāk kā **{claimable}**.",
        "calc_note": "Aprēķins: vienkāršie procenti (bez procentu procentiem). Informatīvs aprēķins, nav juridiska konsultācija. Ja izīrētājs ir uzņēmējs vai līgums ir patērētāju līgums, var būt papildu ierobežojumi.",
        # --- čats
        "chat_title": "💬 Čats ar līgumu",
        "chat_greeting": "Sveiks. Kādi tev radās jautājumi par šo līgumu?",
        "chat_placeholder": "Uzdodiet jautājumu par līgumu...",
        "chat_hide_help": "Paslēpt (saruna saglabāsies)",
        "chat_open_help": "Atvērt čatu",
        "chat_cap_mock": "🧪 Testa režīms: iebūvētas atbildes, Gemini netiek izsaukts (0.00$).",
        "chat_cap_real": "Jautājums tiek sūtīts Gemini kopā ar līguma tekstu un analīzi.",
        "chat_spinner": "💬 Gemini sagatavo atbildi...",
        "chat_no_key": "⚠️ Ievadiet Gemini API atslēgu sānu joslā vai ieslēdziet testa režīmu.",
        "chat_error": "❌ Neizdevās sagatavot atbildi: {error}",
        "mock_pets": "🧪 **Mājdzīvnieki:** šajā līgumā (6. sadaļa) nav atrunāti noteikumi par mājdzīvnieku turēšanu. Tāpēc nav skaidrs, vai suni vai kaķi drīkst turēt, kas atbild par bojājumiem un kā tiek risināti trokšņa vai tīrības strīdi. Ieteikums: pievienot punktu, piemēram, „Īrnieks drīkst turēt vienu mājdzīvnieku ar izīrētāja rakstveida piekrišanu un atbild par tā radītajiem bojājumiem.” Vispārīgie lietošanas noteikumi: {dtil17}.",
        "mock_terminate": "🧪 **Uzteikums:** īrniekam ir tiesības vienpusēji atkāpties no līguma, rakstveidā brīdinot izīrētāju vismaz 1 mēnesi iepriekš (līgumiem līdz 10 gadiem) {dtil27} {dtil20}. Izīrētājs var izbeigt līgumu tikai likumā noteiktajos gadījumos {dtil21}, piemēram, ja īre kavēta vairāk nekā 2 mēnešus {dtil24}. Pārbaudiet, vai līguma uzteikuma punkts atbilst šiem termiņiem.",
        "mock_deposit": "🧪 **Drošības nauda:** tā nedrīkst pārsniegt divu mēnešu īres maksu un jāatdod ne vēlāk kā dienā, kad telpa tiek atbrīvota, ja līgumā nav noteikts citādi {dtil12}. Līguma 4.3. punkts ļauj ieturēt naudu bez saskaņošanas, tāpēc ieteicams prasīt abpusēji parakstītu defektu aktu.",
        "mock_land": "🧪 **Zemesgrāmata:** ja telpu pārdod, jaunajam īpašniekam līgums ir saistošs tikai tad, ja īres tiesības ir ierakstītas zemesgrāmatā {dtil28} {dtil35}. Pretējā gadījumā viņš var prasīt telpas atbrīvot 2 mēnešu laikā.",
        "mock_penalty": "🧪 **Līgumsods un procenti:** līgumsods par kavējumu kopumā nevar pārsniegt 10% no parāda {cl1716}, un tam jābūt samērīgam {cl1717}. Konkrētu summu aprēķiniet cilnē „🧮 Soda kalkulators”.",
        "mock_rent": "🧪 **Īres maksa:** līguma laikā to drīkst mainīt tikai tad, ja līgumā ir paredzēti maksas grozīšanas nosacījumi un kārtība {dtil10}.",
        "mock_repairs": "🧪 **Remonti:** līguma 6.1.2. punktā „sīkie remontdarbi” nav definēti, tāpēc var rasties strīds, kurš maksā par sadzīves tehnikas vai santehnikas bojājumiem. Izīrētājs savukārt nodrošina telpas lietošanas iespēju visu līguma laiku {dtil18}.",
        "mock_default": "🧪 **Testa atbilde:** jautājums saņemts („{question}”). Testa režīmā Gemini netiek izsaukts. Īstajā režīmā atbildi sagatavotu Gemini, balstoties uz līguma tekstu. Pamēģiniet jautāt par mājdzīvniekiem, uzteikumu, drošības naudu, zemesgrāmatu, līgumsodu, īres maksu vai remontiem.",
    },
    "English": {
        "reset_btn": "🔄 Clear session / Start over",
        "sidebar_toggle": "Show / hide the settings panel",
        "tab_risks": "📊 Risk analysis",
        "tab_calc": "🧮 Penalty calculator",
        "tab_docs": "📄 Document download",
        "tab_chat": "💬 Chat with the contract",
        "tab_fixes": "### 🟢 Recommendations and corrections",
        "cite_cl_gen": "Civil Law, Art. {n}",
        "cite_cl_acc": "Civil Law, Art. {n}",
        "calc_title": "🧮 Penalty and interest calculator",
        "calc_autofill": "Fields were filled in automatically from the contract text (where found). Please check them and correct if needed.",
        "calc_kind": "Type of sanction",
        "calc_kind_daily": "Contractual penalty / late fee (% per day)",
        "calc_kind_annual": "Late-payment interest (% per year)",
        "calc_principal": "Debt amount / rent (EUR)",
        "calc_days": "Days overdue",
        "calc_rate_day": "Rate in the contract (% per day)",
        "calc_rate_year": "Rate in the contract (% per year)",
        "calc_cap": "The contract sets an overall cap on the penalty (no more than 10% of the debt)",
        "calc_enter": "Enter the amount, the number of days and the rate to see the calculation.",
        "calc_m_contract": "Under the contract terms",
        "calc_m_cap": "Legal maximum (10%)",
        "calc_m_diff": "Difference vs. statutory interest",
        "calc_m_stat": "Statutory interest (6% per year)",
        "calc_t_error": "🔴 Exceeds the legal limit",
        "calc_t_warn": "🟡 Attention",
        "calc_t_ok": "🟢 Complies with the rules",
        "calc_t_risk": "🔴 Risk",
        "calc_msg_exceed": "The penalty calculated under the contract ({total}) exceeds 10% of the debt ({cap}). Under {cl1716_gen} (part 3), a penalty for late or improper performance cannot in total exceed 10% of the principal debt. The excess ({excess}) cannot be claimed.",
        "calc_msg_nocap": "The contract sets no overall cap on the penalty. At {rate}% per day the 10% limit will be reached after {days_to_cap} days. Under {cl1716_gen} the contract must state that the penalty in total does not exceed 10%.",
        "calc_msg_highrate": "A rate of {rate}% per day equals about {annual}% per year, i.e. {multiple} times the statutory interest (6% per year, {cl1765}). The law sets no specific daily rate limit, but a court reduces an excessive penalty ({cl1717}, {cl1724}). 0.1% per day is a more commonly seen proportionate level (an informal benchmark, not a legal rule).",
        "calc_msg_ok_daily": "The penalty ({total}) does not exceed the 10% cap ({cap}) under {cl1716_gen}, the contract sets the cap, and the rate does not exceed 0.1% per day.",
        "calc_msg_ok_annual": "The rate ({rate}% per year) does not exceed the statutory interest (6% per year, {cl1765}).",
        "calc_msg_warn_annual": "The rate ({rate}% per year) is {multiple} times higher than the statutory interest (6% per year, {cl1765}). The Civil Law sets no specific upper limit, but interest that is disproportionate and contrary to fair dealing is unlawful ({cl1764}, {cl1717}).",
        "calc_msg_risk_annual": "The rate ({rate}% per year, about {rate_day}% per day) is very high ({multiple} times the statutory interest) and may be found disproportionate and unlawful ({cl1764}, {cl1717}).",
        "calc_both": "**If the landlord claims both statutory interest and a contractual penalty** (under {cl1722_acc}): the penalty can be claimed only to the extent that it exceeds the interest claimed ({statutory}). The penalty in this case: no more than **{claimable}**.",
        "calc_note": "Calculation: simple interest (no compounding). Informational calculation, not legal advice. If the landlord is a business or the contract is a consumer contract, additional restrictions may apply.",
        "chat_title": "💬 Chat about the contract",
        "chat_greeting": "Hello. What questions do you have about this contract?",
        "chat_placeholder": "Ask a question about the contract...",
        "chat_hide_help": "Hide (the conversation is kept)",
        "chat_open_help": "Open chat",
        "chat_cap_mock": "🧪 Test mode: built-in answers, Gemini is not called ($0.00).",
        "chat_cap_real": "Your question is sent to Gemini together with the contract text and the analysis.",
        "chat_spinner": "💬 Gemini is preparing an answer...",
        "chat_no_key": "⚠️ Enter your Gemini API key in the sidebar or turn on test mode.",
        "chat_error": "❌ Could not prepare an answer: {error}",
        "mock_pets": "🧪 **Pets:** this contract (section 6) has no rules on keeping pets. It is therefore unclear whether a dog or a cat is allowed, who is responsible for damage and how noise or cleanliness disputes are resolved. Suggestion: add a clause such as “The Tenant may keep one pet with the Landlord’s written consent and is liable for any damage it causes.” General rules on use of the premises: {dtil17}.",
        "mock_terminate": "🧪 **Termination:** the tenant may withdraw from the contract unilaterally by giving the landlord at least 1 month’s written notice (for contracts of up to 10 years) {dtil27} {dtil20}. The landlord may terminate the contract only in the cases set out in the law {dtil21}, for example if rent is more than 2 months overdue {dtil24}. Check that the termination clause of the contract matches these periods.",
        "mock_deposit": "🧪 **Security deposit:** it may not exceed two months’ rent and must be returned no later than the day the premises are vacated, unless the contract says otherwise {dtil12}. Clause 4.3 of the contract allows money to be withheld without agreement, so it is advisable to require a defects report signed by both parties.",
        "mock_land": "🧪 **Land register:** if the premises are sold, the contract binds the new owner only if the lease rights are entered in the land register {dtil28} {dtil35}. Otherwise the new owner may demand that the premises be vacated within 2 months.",
        "mock_penalty": "🧪 **Penalty and interest:** a penalty for late payment cannot in total exceed 10% of the debt {cl1716} and must be proportionate {cl1717}. Calculate the exact amount with the “🧮 Penalty calculator” tab.",
        "mock_rent": "🧪 **Rent:** during the contract it may be changed only if the contract sets out conditions and a procedure for changing it {dtil10}.",
        "mock_repairs": "🧪 **Repairs:** clause 6.1.2 of the contract does not define “minor repairs”, so a dispute may arise over who pays for damage to appliances or plumbing. The landlord, in turn, must keep the premises usable throughout the contract {dtil18}.",
        "mock_default": "🧪 **Test answer:** your question was received (“{question}”). In test mode Gemini is not called. In real mode Gemini would answer based on the contract text. Try asking about pets, termination, the security deposit, the land register, penalties, rent or repairs.",
    },
    "Русский": {
        "reset_btn": "🔄 Очистить сессию / Начать заново",
        "sidebar_toggle": "Показать / скрыть панель настроек",
        "tab_risks": "📊 Анализ рисков",
        "tab_calc": "🧮 Калькулятор штрафов",
        "tab_docs": "📄 Скачать документы",
        "tab_chat": "💬 Чат с договором",
        "tab_fixes": "### 🟢 Рекомендации и исправления",
        "cite_cl_gen": "Гражданский закон, ст. {n}",
        "cite_cl_acc": "Гражданский закон, ст. {n}",
        "calc_title": "🧮 Калькулятор штрафов и процентов",
        "calc_autofill": "Поля заполнены автоматически из текста договора (если найдено). Проверьте и при необходимости исправьте.",
        "calc_kind": "Вид санкции",
        "calc_kind_daily": "Договорная неустойка / пеня (% в день)",
        "calc_kind_annual": "Проценты за просрочку (% годовых)",
        "calc_principal": "Сумма долга / арендная плата (EUR)",
        "calc_days": "Дней просрочки",
        "calc_rate_day": "Ставка в договоре (% в день)",
        "calc_rate_year": "Ставка в договоре (% годовых)",
        "calc_cap": "В договоре установлен общий предел неустойки (не более 10% от долга)",
        "calc_enter": "Введите сумму, количество дней и ставку, чтобы увидеть расчёт.",
        "calc_m_contract": "По условиям договора",
        "calc_m_cap": "Максимум по закону (10%)",
        "calc_m_diff": "Разница с законными процентами",
        "calc_m_stat": "Законные проценты (6% годовых)",
        "calc_t_error": "🔴 Превышает установленное законом",
        "calc_t_warn": "🟡 Внимание",
        "calc_t_ok": "🟢 Соответствует нормам",
        "calc_t_risk": "🔴 Риск",
        "calc_msg_exceed": "Рассчитанная по договору неустойка ({total}) превышает 10% от долга ({cap}). Согласно {cl1716_gen} (часть 3), неустойка за ненадлежащее исполнение (просрочку) в совокупности не может превышать 10% от основного долга. Превышающая часть ({excess}) не подлежит взысканию.",
        "calc_msg_nocap": "В договоре не установлен общий предел неустойки. При ставке {rate}% в день порог в 10% будет достигнут через {days_to_cap} дн. Согласно {cl1716_gen} в договоре должно быть указано, что неустойка в совокупности не превышает 10%.",
        "calc_msg_highrate": "Ставка {rate}% в день соответствует ~{annual}% годовых, то есть в {multiple} раза выше законных процентов (6% годовых, {cl1765}). Закон не устанавливает конкретного предела дневной ставки, но суд снижает чрезмерную неустойку ({cl1717}, {cl1724}). 0,1% в день — чаще встречающийся соразмерный уровень (ориентир, не норма закона).",
        "calc_msg_ok_daily": "Неустойка ({total}) не превышает предел в 10% ({cap}) согласно {cl1716_gen}, предел в договоре установлен, а ставка не превышает 0,1% в день.",
        "calc_msg_ok_annual": "Ставка ({rate}% годовых) не превышает законные проценты (6% годовых, {cl1765}).",
        "calc_msg_warn_annual": "Ставка ({rate}% годовых) в {multiple} раза выше законных процентов (6% годовых, {cl1765}). Гражданский закон не устанавливает конкретного верхнего предела, но несоразмерные и противоречащие добросовестной практике проценты противоправны ({cl1764}, {cl1717}).",
        "calc_msg_risk_annual": "Ставка ({rate}% годовых, ~{rate_day}% в день) очень высока (в {multiple} раза выше законных процентов) и может быть признана несоразмерной и противоправной ({cl1764}, {cl1717}).",
        "calc_both": "**Если арендодатель требует и законные проценты, и неустойку** (согласно {cl1722_acc}): неустойку можно требовать только в части, превышающей заявленную сумму процентов ({statutory}). Неустойка в этом случае: не более **{claimable}**.",
        "calc_note": "Расчёт: простые проценты (без начисления процентов на проценты). Информационный расчёт, не юридическая консультация. Если арендодатель — предприниматель или договор является потребительским, могут действовать дополнительные ограничения.",
        "chat_title": "💬 Чат о договоре",
        "chat_greeting": "Здравствуйте. Какие у вас возникли вопросы по этому договору?",
        "chat_placeholder": "Задайте вопрос о договоре...",
        "chat_hide_help": "Скрыть (разговор сохранится)",
        "chat_open_help": "Открыть чат",
        "chat_cap_mock": "🧪 Тестовый режим: встроенные ответы, Gemini не вызывается ($0.00).",
        "chat_cap_real": "Вопрос отправляется в Gemini вместе с текстом договора и анализом.",
        "chat_spinner": "💬 Gemini готовит ответ...",
        "chat_no_key": "⚠️ Введите API-ключ Gemini в боковой панели или включите тестовый режим.",
        "chat_error": "❌ Не удалось подготовить ответ: {error}",
        "mock_pets": "🧪 **Домашние животные:** в этом договоре (раздел 6) нет правил содержания домашних животных. Поэтому неясно, можно ли держать собаку или кошку, кто отвечает за повреждения и как решаются споры о шуме или чистоте. Рекомендация: добавить пункт, например: «Арендатор может держать одно домашнее животное с письменного согласия арендодателя и отвечает за причинённый им ущерб». Общие правила пользования помещением: {dtil17}.",
        "mock_terminate": "🧪 **Расторжение:** арендатор вправе в одностороннем порядке отказаться от договора, письменно уведомив арендодателя не менее чем за 1 месяц (для договоров до 10 лет) {dtil27} {dtil20}. Арендодатель может расторгнуть договор только в случаях, предусмотренных законом {dtil21}, например, если просрочка арендной платы превышает 2 месяца {dtil24}. Проверьте, соответствует ли пункт договора о расторжении этим срокам.",
        "mock_deposit": "🧪 **Залог:** он не может превышать арендную плату за два месяца и должен быть возвращён не позднее дня освобождения помещения, если в договоре не предусмотрено иное {dtil12}. Пункт 4.3 договора позволяет удерживать деньги без согласования, поэтому желательно требовать акт о дефектах, подписанный обеими сторонами.",
        "mock_land": "🧪 **Земельная книга:** если помещение продаётся, договор обязателен для нового собственника только в том случае, если права аренды внесены в земельную книгу {dtil28} {dtil35}. В противном случае он может потребовать освободить помещение в течение 2 месяцев.",
        "mock_penalty": "🧪 **Неустойка и проценты:** неустойка за просрочку в совокупности не может превышать 10% от долга {cl1716} и должна быть соразмерной {cl1717}. Точную сумму рассчитайте во вкладке «🧮 Калькулятор штрафов».",
        "mock_rent": "🧪 **Арендная плата:** в течение срока договора её можно изменять только в том случае, если в договоре предусмотрены условия и порядок изменения платы {dtil10}.",
        "mock_repairs": "🧪 **Ремонт:** в пункте 6.1.2 договора «мелкий ремонт» не определён, поэтому может возникнуть спор о том, кто платит за повреждения бытовой техники или сантехники. Арендодатель, в свою очередь, обеспечивает возможность пользоваться помещением в течение всего срока договора {dtil18}.",
        "mock_default": "🧪 **Тестовый ответ:** вопрос получен («{question}»). В тестовом режиме Gemini не вызывается. В реальном режиме ответ подготовил бы Gemini на основе текста договора. Попробуйте спросить о домашних животных, расторжении, залоге, земельной книге, неустойке, арендной плате или ремонте.",
    },
}
for _l, _extra in TR_NEW.items():
    TR[_l].update(_extra)

# Labotā līguma lejupielāde
TR_FIX = {
    "Latviešu": {
        "fx_header": "### ✍️ Labotais līgums",
        "fx_intro": "Atzīmējiet labojumus, kurus ievietot līgumā. Labotajā līgumā jaunais teksts ir sarkanā krāsā, aizstātais teksts ir pārsvītrots.",
        "fx_text_label": "Līguma teksts (var labot vai ielīmēt, ja līgums augšupielādēts kā PDF vai attēls):",
        "fx_no_text": "Līguma teksts nav pieejams (PDF vai attēls). Ielīmējiet līguma tekstu laukā augstāk, lai izveidotu labotu līgumu.",
        "fx_no_fixes": "Analīzē nav atrasts neviens ieteiktais labojums ar piedāvātu formulējumu.",
        "fx_none_selected": "Nav atzīmēts neviens labojums.",
        "fx_all": "Atzīmēt visus", "fx_clear": "Noņemt visus",
        "fx_dl_word": "📝 Labotais līgums Word (.docx)", "fx_dl_pdf": "📝 Labotais līgums PDF (.pdf)",
        "fx_title": "LABOTAIS ĪRES LĪGUMS",
        "fx_legend": "Ievietotie labojumi ir sarkanā krāsā; aizstātais teksts ir pārsvītrots.",
        "fx_extra": "PAPILDINĀJUMI (punkti, kurus neizdevās automātiski atrast līguma tekstā)",
        "fx_count": "Atzīmēti labojumi: {n} no {m}",
    },
    "English": {
        "fx_header": "### ✍️ Corrected contract",
        "fx_intro": "Tick the corrections to insert into the contract. In the corrected contract new text is red and replaced text is struck through.",
        "fx_text_label": "Contract text (you can edit it, or paste it if the contract was uploaded as a PDF or image):",
        "fx_no_text": "The contract text is not available (PDF or image). Paste the contract text in the field above to create a corrected contract.",
        "fx_no_fixes": "No recommended correction with a proposed wording was found in the analysis.",
        "fx_none_selected": "No correction is selected.",
        "fx_all": "Select all", "fx_clear": "Clear all",
        "fx_dl_word": "📝 Corrected contract Word (.docx)", "fx_dl_pdf": "📝 Corrected contract PDF (.pdf)",
        "fx_title": "CORRECTED LEASE CONTRACT",
        "fx_legend": "Inserted corrections are in red; replaced text is struck through.",
        "fx_extra": "ADDITIONS (points that could not be found automatically in the contract text)",
        "fx_count": "Selected corrections: {n} of {m}",
    },
    "Русский": {
        "fx_header": "### ✍️ Исправленный договор",
        "fx_intro": "Отметьте исправления, которые нужно внести в договор. В исправленном договоре новый текст выделен красным, заменённый текст зачёркнут.",
        "fx_text_label": "Текст договора (можно править или вставить, если договор загружен как PDF или изображение):",
        "fx_no_text": "Текст договора недоступен (PDF или изображение). Вставьте текст договора в поле выше, чтобы создать исправленный договор.",
        "fx_no_fixes": "В анализе не найдено рекомендованных исправлений с предлагаемой формулировкой.",
        "fx_none_selected": "Не выбрано ни одного исправления.",
        "fx_all": "Отметить все", "fx_clear": "Снять все",
        "fx_dl_word": "📝 Исправленный договор Word (.docx)", "fx_dl_pdf": "📝 Исправленный договор PDF (.pdf)",
        "fx_title": "ИСПРАВЛЕННЫЙ ДОГОВОР АРЕНДЫ",
        "fx_legend": "Внесённые исправления выделены красным; заменённый текст зачёркнут.",
        "fx_extra": "ДОПОЛНЕНИЯ (пункты, которые не удалось автоматически найти в тексте договора)",
        "fx_count": "Отмечено исправлений: {n} из {m}",
    },
}
for _l, _extra in TR_FIX.items():
    if _l in TR:
        TR[_l].update(_extra)

# PWA: pamācība, kā pievienot lietotni sākuma ekrānam (sānu joslas apakšā)
TR_PWA = {
    "Latviešu": {
        "pwa_btn": "📲 Instalēt lietotni",
        "pwa_installed": "✅ Lietotne ir instalēta",
        "pwa_pop_title": "Instalējiet lietotni",
        "pwa_pop_text": "Atveriet līgumu analizētāju ar vienu pieskārienu no tālruņa vai planšetdatora sākuma ekrāna.",
        "pwa_close": "Aizvērt",
        "pwa_ios_h": "🍎 iPhone / iPad (Safari)",
        "pwa_ios": "1) Nospiediet **„Kopīgot“** (kvadrāts ar bultiņu ⬆️) ekrāna apakšā (iPad – augšā; jaunākajā iOS tas ir izvēlnē **⋯**).<br>2) Izvēlieties **„Pievienot sākuma ekrānam“** (Add to Home Screen).<br>3) Nospiediet **„Pievienot“**.",
        "pwa_ios_chrome_h": "🍎 iPhone / iPad (Chrome)",
        "pwa_ios_chrome": "1) Nospiediet **„Kopīgot“** (⬆️) adreses joslas labajā pusē.<br>2) Izvēlieties **„Pievienot sākuma ekrānam“**.<br>3) Nospiediet **„Pievienot“**.",
        "pwa_mac_h": "🍎 Mac (Safari)",
        "pwa_mac": "Augšējā izvēlnē **„Fails“** (File) izvēlieties **„Pievienot Dokam“** (Add to Dock) un nospiediet **„Pievienot“**.",
        "pwa_android_h": "🤖 Android (Chrome)",
        "pwa_android": "Nospiediet **trīs punktiņus** (⋮) augšējā stūrī un izvēlieties **„Instalēt lietotni“** (Install app).",
    },
    "English": {
        "pwa_btn": "📲 Install app",
        "pwa_installed": "✅ App is installed",
        "pwa_pop_title": "Install the app",
        "pwa_pop_text": "Open the contract analyzer with one tap from the home screen of your phone or tablet.",
        "pwa_close": "Close",
        "pwa_ios_h": "🍎 iPhone / iPad (Safari)",
        "pwa_ios": "1) Tap **Share** (square with arrow ⬆️) at the bottom of the screen (iPad – at the top; on the newest iOS it is in the **⋯** menu).<br>2) Choose **Add to Home Screen**.<br>3) Tap **Add**.",
        "pwa_ios_chrome_h": "🍎 iPhone / iPad (Chrome)",
        "pwa_ios_chrome": "1) Tap **Share** (⬆️) on the right side of the address bar.<br>2) Choose **Add to Home Screen**.<br>3) Tap **Add**.",
        "pwa_mac_h": "🍎 Mac (Safari)",
        "pwa_mac": "In the top menu **File** choose **Add to Dock**, then click **Add**.",
        "pwa_android_h": "🤖 Android (Chrome)",
        "pwa_android": "Tap the **three dots** (⋮) in the top corner and choose **Install app**.",
    },
    "Русский": {
        "pwa_btn": "📲 Установить приложение",
        "pwa_installed": "✅ Приложение установлено",
        "pwa_pop_title": "Установите приложение",
        "pwa_pop_text": "Открывайте анализатор договоров одним касанием с главного экрана телефона или планшета.",
        "pwa_close": "Закрыть",
        "pwa_ios_h": "🍎 iPhone / iPad (Safari)",
        "pwa_ios": "1) Нажмите **«Поделиться»** (квадрат со стрелкой ⬆️) внизу экрана (iPad – вверху; в новейшей iOS – в меню **⋯**).<br>2) Выберите **«На экран „Домой“»** (Add to Home Screen).<br>3) Нажмите **«Добавить»**.",
        "pwa_ios_chrome_h": "🍎 iPhone / iPad (Chrome)",
        "pwa_ios_chrome": "1) Нажмите **«Поделиться»** (⬆️) справа в адресной строке.<br>2) Выберите **«На экран „Домой“»**.<br>3) Нажмите **«Добавить»**.",
        "pwa_mac_h": "🍎 Mac (Safari)",
        "pwa_mac": "В верхнем меню **«Файл»** (File) выберите **«Добавить в Dock»** (Add to Dock) и нажмите **«Добавить»**.",
        "pwa_android_h": "🤖 Android (Chrome)",
        "pwa_android": "Нажмите **три точки** (⋮) в верхнем углу и выберите **«Установить приложение»** (Install app).",
    },
}
for _l, _extra in TR_PWA.items():
    if _l in TR:
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


# ---------------------------------------------------------------
# Modelis: tikai gemini-3.8-flash (rezerves modeļu nav)
# ---------------------------------------------------------------
MODEL_PRIMARY = "gemini-3.8-flash"
# Cena USD par 1 miljonu žetonu (ievade, izvade; izvade ietver thinking). Kešotā ievade: 0.075.
PRICE_PER_M = (0.75, 3.75)


def cost_usd(tokens_in, tokens_out):
    return (tokens_in * PRICE_PER_M[0] + tokens_out * PRICE_PER_M[1]) / 1_000_000


class GeminiModel:
    """Ap GenerativeModel: vienmēr izmanto tikai gemini-3.8-flash."""

    def __init__(self, generation_config=None):
        self.cfg = generation_config

    def generate_content(self, contents, **kw):
        return genai.GenerativeModel(MODEL_PRIMARY, generation_config=self.cfg).generate_content(contents, **kw)


@st.cache_data(show_spinner=False)
def translate_ui(lang_name, _api_key):
    """Iztulko visus saskarnes tekstus izvēlētajā valodā (rezultāts tiek saglabāts kešatmiņā)."""
    genai.configure(api_key=_api_key)
    model = GeminiModel(genai.GenerationConfig(temperature=0, response_mime_type="application/json"))
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

# Tēmas pārslēgšana: 🔵 tumši zilais (pamata režīms), ☀️ baltais, ⚫ melnais (pilnīgi melns #000, OLED taupa enerģiju).
# Tumši zilais nāk no .streamlit/config.toml (bez pārklājuma), pārējie divi tiek uzlikti ar CSS.
# Pārslēdz pārlūkā (JavaScript), tāpēc Streamlit netiek pārlādēts un analīze netiek pārtraukta.
# Izvēle tiek atcerēta pārlūkā.
import streamlit.components.v1 as components

THEME_HTML = """
<style>
  body { margin: 0; font-family: sans-serif; }
  .row { display: flex; gap: 6px; }
  button { flex: 1; padding: 7px 0; font-size: 17px; border-radius: 8px;
           border: 1px solid #64748b; background: transparent; cursor: pointer; opacity: 0.75; }
  button.on { border: 2px solid #4f7df0; background: rgba(79,125,240,0.25); opacity: 1; }
</style>
<div class="row">
  <button data-t="navy" title="Dark blue">🔵</button>
  <button data-t="light" title="Light">☀️</button>
  <button data-t="oled" title="Black (OLED)">⚫</button>
</div>
<script>
(function () {
  var P = window.parent, D = P.document;
  // fons, sānu josla, lauki, teksts, apmales, akcents, galvenā poga
  var THEMES = {
    navy:  null,   // config.toml krāsas, nekas netiek pārklāts
    light: { bg: "#ffffff", bg2: "#eef2fa", fld: "#ffffff", fg: "#0f172a", bd: "#cbd5e1", ln: "#1e3a8a", btn: "#1e3a8a", alert: "#f1f5fb" },
    oled:  { bg: "#000000", bg2: "#000000", fld: "#000000", fg: "#e5e7eb", bd: "#262b36", ln: "#7da2ff", btn: "#1e3a8a", alert: "#0a0d14" }
  };
  function css(c) {
    return "" +
      ".stApp, [data-testid='stHeader'], [data-testid='stBottom'] > div, [data-testid='stBottomBlockContainer'] { background: " + c.bg + " !important; }" +
      "[data-testid='stSidebar'], [data-testid='stSidebar'] > div { background: " + c.bg2 + " !important; }" +
      ".stApp, .stApp p, .stApp li, .stApp label, .stApp span, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp small, .stApp td, .stApp th, .stApp summary { color: " + c.fg + " !important; }" +
      ".stApp button p, .stApp button span { color: inherit !important; }" +
      ".stApp a, .stApp a span { color: " + c.ln + " !important; }" +
      // galvenā poga (tumši zila abos režīmos)
      "[data-testid='stBaseButton-primary'] { background: " + c.btn + " !important; border-color: " + c.btn + " !important; color: #fff !important; }" +
      "[data-testid='stBaseButton-primary']:hover { filter: brightness(1.15); }" +
      // pārējās pogas
      "[data-testid='stBaseButton-secondary'], [data-testid='stBaseButton-secondaryFormSubmit'], .stDownloadButton button, [data-testid='stPopover'] button" +
      " { background: " + c.fld + " !important; color: " + c.fg + " !important; border-color: " + c.bd + " !important; }" +
      "[data-testid='stBaseButton-secondary']:hover, .stDownloadButton button:hover { border-color: " + c.ln + " !important; color: " + c.ln + " !important; }" +
      // ievades lauki
      ".stTextInput input, .stTextArea textarea, .stNumberInput input, [data-baseweb='input'], [data-baseweb='base-input'], [data-baseweb='textarea']," +
      " [data-baseweb='select'] > div, [data-testid='stFileUploaderDropzone'], [data-testid='stChatInput'], [data-testid='stChatInput'] > div," +
      " [data-testid='stChatInput'] textarea { background: " + c.fld + " !important; color: " + c.fg + " !important; border-color: " + c.bd + " !important; }" +
      ".stNumberInput button { background: " + c.fld + " !important; color: " + c.fg + " !important; border-color: " + c.bd + " !important; }" +
      "[data-testid='stFileUploaderDropzone'] button { background: " + c.bg2 + " !important; color: " + c.fg + " !important; }" +
      // cilnes
      "button[role='tab'] p { color: " + c.fg + " !important; }" +
      "button[role='tab'][aria-selected='true'] p { color: " + c.ln + " !important; font-weight: 700; }" +
      "[data-baseweb='tab-highlight'] { background-color: " + c.ln + " !important; }" +
      "[data-baseweb='tab-border'] { background-color: " + c.bd + " !important; }" +
      // sadaļas, ziņojumi, čats, metrikas
      "[data-testid='stExpander'] details, [data-testid='stExpander'] summary { background: " + c.bg + " !important; border-color: " + c.bd + " !important; }" +
      "[data-testid='stAlert'] > div { background: " + c.alert + " !important; border: 1px solid " + c.bd + " !important; }" +
      "[data-testid='stAlert'] * { color: " + c.fg + " !important; }" +
      "[data-testid='stChatMessage'] { background: transparent !important; }" +
      "[data-testid='stVerticalBlockBorderWrapper'] { border-color: " + c.bd + " !important; }" +
      "[data-testid='stMetricValue'] *, [data-testid='stMetricLabel'] * { color: " + c.fg + " !important; }" +
      "[data-baseweb='popover'] > div, [data-baseweb='popover'] [data-testid='stVerticalBlock'], [data-baseweb='popover'] ul, [data-baseweb='menu']" +
      " { background: " + c.bg2 + " !important; border-color: " + c.bd + " !important; }" +
      "[data-baseweb='popover'] p, [data-baseweb='popover'] li, [data-baseweb='popover'] small { color: " + c.fg + " !important; }" +
      "[data-baseweb='checkbox'] span, [data-baseweb='radio'] div:first-child { border-color: #64748b !important; }" +
      "hr { border-color: " + c.bd + " !important; }" +
      "code, pre { background: " + c.alert + " !important; color: " + c.fg + " !important; }";
  }
  function apply(t) {
    if (!(t in THEMES)) t = "navy";
    var s = D.getElementById("app-theme-style");
    if (!s) { s = D.createElement("style"); s.id = "app-theme-style"; D.head.appendChild(s); }
    s.textContent = THEMES[t] ? css(THEMES[t]) : "";
    D.documentElement.setAttribute("data-app-theme", t);
    try { P.localStorage.setItem("appTheme3", t); } catch (e) {}
    var bs = document.querySelectorAll("button");
    for (var i = 0; i < bs.length; i++) { bs[i].classList.toggle("on", bs[i].getAttribute("data-t") === t); }
  }
  var saved = "navy";
  try { saved = P.localStorage.getItem("appTheme3") || "navy"; } catch (e) {}
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

# ---------------------------------------------------------------
# Bīdāms sānu panelis (drawer): pastāvīga poga kreisajā augšējā stūrī atver/aizver
# st.sidebar. Izmanto Streamlit iebūvētās pogas, tāpēc stāvoklis paliek sinhronizēts.
# Kad panelis aizvērts, galvenais saturs kļūst platāks. Izvēle tiek atcerēta pārlūkā.
# Īsinājumtaustiņš: Alt+S. Mobilajā ierīcē panelis aizveras ar Esc vai pieskārienu ārpus tā.
# ---------------------------------------------------------------
DRAWER_HTML = """
<script>
(function () {
  var P = window.parent, D = P.document;
  var TITLE = __TITLE__;
  function sidebar() { return D.querySelector('[data-testid="stSidebar"]'); }
  function isOpen() {
    // Streamlit 1.4x+: data-collapsed="true|false"; vecākās versijās: aria-expanded
    var sb = sidebar(); if (!sb) return false;
    var dc = sb.getAttribute("data-collapsed");
    if (dc !== null) return dc !== "true";
    return sb.getAttribute("aria-expanded") !== "false";
  }
  function clickNative(open) {
    var b = open
      ? D.querySelector('[data-testid="stExpandSidebarButton"]') ||
        D.querySelector('[data-testid="stSidebarCollapsedControl"] button')
      : D.querySelector('[data-testid="stSidebarCollapseButton"] button') ||
        D.querySelector('[data-testid="stSidebarCollapseButton"]');
    if (b) { b.click(); return true; }
    return false;
  }
  function setOpen(open, remember) {
    if (isOpen() !== open) clickNative(open);
    if (remember) { try { P.localStorage.setItem("appSidebarOpen", open ? "1" : "0"); } catch (e) {} }
  }
  function toggle() { setOpen(!isOpen(), true); }

  // Stils: poga, platāks saturs, kad panelis aizvērts
  var st = D.getElementById("app-drawer-style");
  if (!st) { st = D.createElement("style"); st.id = "app-drawer-style"; D.head.appendChild(st); }
  st.textContent =
    // Neliela, bet izteikta poga kreisajā augšējā stūrī; seko paneļa malai
    "#app-drawer-btn { position: fixed; top: calc(10px + env(safe-area-inset-top, 0px)); left: 10px; z-index: 1000100;" +
    " width: 34px; height: 34px; border: 2px solid rgba(255,255,255,0.85); border-radius: 10px;" +
    " background: #1e3a8a; color: #fff; font-size: 17px; font-weight: 900; line-height: 1;" +
    " display: flex; align-items: center; justify-content: center;" +
    " cursor: pointer; padding: 0; box-shadow: 0 2px 10px rgba(0,0,0,0.35);" +
    " transition: left 0.3s ease, background 0.2s, transform 0.15s; }" +
    "#app-drawer-btn:hover { background: #172f70; transform: scale(1.08); }" +
    "#app-drawer-btn:focus-visible { outline: 3px solid #93c5fd; outline-offset: 2px; }" +
    // Streamlit iebūvētās pogas paslēptas (tās joprojām tiek izmantotas no JavaScript)
    "[data-testid='stExpandSidebarButton'], [data-testid='stSidebarCollapsedControl']," +
    " [data-testid='stSidebarCollapseButton'] { opacity: 0 !important; pointer-events: none !important; }" +
    ".stApp:has([data-testid='stSidebar'][data-collapsed='true']) [data-testid='stMainBlockContainer']," +
    " .stApp:has([data-testid='stSidebar'][aria-expanded='false']) [data-testid='stMainBlockContainer']" +
    " { max-width: min(1200px, 100%) !important; transition: max-width 0.3s ease; }" +
    "[data-testid='stElementContainer']:has(iframe[height='0']) { display: none !important; }" +
    "@media print { #app-drawer-btn { display: none !important; } }";

  var btn = D.getElementById("app-drawer-btn");
  if (!btn) {
    btn = D.createElement("button"); btn.id = "app-drawer-btn"; btn.type = "button";
    D.body.appendChild(btn);
  }
  btn.onclick = toggle;
  btn.title = TITLE + " (Alt+S)";
  btn.setAttribute("aria-label", TITLE);

  function sync() {
    var open = isOpen(), sb = sidebar();
    btn.textContent = open ? "\u276E" : "\u276F";      // ❮ / ❯
    btn.setAttribute("aria-expanded", open ? "true" : "false");
    btn.style.left = (open && sb) ? Math.max(10, sb.getBoundingClientRect().right + 10) + "px" : "10px";
  }

  if (!P.__appDrawerInit) {
    P.__appDrawerInit = true;
    // Atjaunot pēdējo izvēli
    var saved = null;
    try { saved = P.localStorage.getItem("appSidebarOpen"); } catch (e) {}
    if (saved === "0" || saved === "1") {
      var tries = 0, want = saved === "1";
      var iv = P.setInterval(function () {
        tries++;
        if (sidebar() && (isOpen() === want || clickNative(want) || tries > 40)) P.clearInterval(iv);
      }, 100);
    }
    P.addEventListener("keydown", function (e) {
      if (e.altKey && (e.code === "KeyS")) { e.preventDefault(); toggle(); }
      else if (e.key === "Escape" && P.innerWidth < 768 && isOpen()) { setOpen(false, false); }
    });
    P.addEventListener("resize", function () { P.requestAnimationFrame(sync); });
    // Sekot panelim (arī tad, ja to aizver/atver ar iebūvētajām pogām vai velkot platumu)
    new P.MutationObserver(function () { P.requestAnimationFrame(sync); })
      .observe(D.body, { attributes: true, subtree: true, attributeFilter: ["data-collapsed", "aria-expanded", "style"] });
    P.setInterval(sync, 500);
  }
  sync();
})();
</script>
"""

components.html(
    DRAWER_HTML.replace("__TITLE__", json.dumps(T.get("sidebar_toggle", "Show / hide the settings panel"))),
    height=0,
)

# ---------------------------------------------------------------
# PWA (Progressive Web App): lietotni var instalēt viedtālrunī / planšetdatorā.
# Streamlit neļauj mainīt lapas <head>, tāpēc saites uz manifestu, iOS meta tagus un
# service worker reģistrāciju ievieto galvenajā lapā ar JavaScript. Faili ir mapē static/
# (pieejami /app/static/..., jo config.toml ir server.enableStaticServing = true).
# ---------------------------------------------------------------
# Šis kods tiek ievietots galvenajā lapā kā <script> (nevis darbojas komponenta iframe),
# tāpēc tas paliek dzīvs arī pēc Streamlit pārlādēm. Tas:
#   1) notver pārlūka "beforeinstallprompt" notikumu (Android Chrome, Edge, datora Chrome),
#      lai instalēšanas logu varētu atvērt ar mūsu pogu;
#   2) lapas atvēršanas brīdī parāda uznirstošo logu ar pogu "Instalēt lietotni" uz 7 sekundēm
#      vai līdz lietotājs kaut kur pieskaras (ja lietotne jau instalēta, logs netiek rādīts);
#   3) iOS (Safari neļauj lapai pašai instalēties) un citos pārlūkos poga parāda īsu norādi.
PWA_JS = r"""
(function () {
  var W = window, D = document;
  function tx() { try { return JSON.parse(D.documentElement.getAttribute("data-pwa-texts") || "{}"); } catch (e) { return {}; } }
  function md(s) {
    return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/&lt;br&gt;/g, "<br>");
  }
  var isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  var UA = navigator.userAgent;
  var isIOSChrome = isIOS && /CriOS/.test(UA);
  // Safari uz Mac (Chrome/Edge uz Mac atbalsta automātisku instalēšanu, Safari – nē)
  var isMacSafari = !isIOS && /Macintosh/.test(UA) && /Safari/.test(UA) && !/Chrome|Chromium|Edg|OPR|Firefox/.test(UA);
  var installed = W.matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
  var deferred = null, box = null, timer = null;

  W.addEventListener("beforeinstallprompt", function (e) { e.preventDefault(); deferred = e; });
  W.addEventListener("appinstalled", function () { installed = true; deferred = null; close(); });

  var css = D.createElement("style");
  css.textContent =
    "#pwa-pop{position:fixed;left:50%;bottom:calc(16px + env(safe-area-inset-bottom,0px));z-index:1000002;" +
    "transform:translate(-50%,20px);opacity:0;transition:opacity .25s,transform .25s;width:min(420px,calc(100vw - 32px));" +
    "box-sizing:border-box;background:#132a5e;color:#e6edfb;border:1px solid #4f7df0;border-radius:14px;" +
    "box-shadow:0 10px 30px rgba(0,0,0,.45);padding:14px 16px 16px;font-family:'Source Sans Pro',sans-serif;overflow:hidden}" +
    "#pwa-pop.on{opacity:1;transform:translate(-50%,0)}" +
    "#pwa-pop .row{display:flex;gap:12px;align-items:center}" +
    "#pwa-pop img{width:48px;height:48px;border-radius:11px;flex:none}" +
    "#pwa-pop h4{margin:0 0 2px;font-size:16px;color:#fff}" +
    "#pwa-pop p{margin:0;font-size:13.5px;line-height:1.4;opacity:.9}" +
    "#pwa-pop .help p{margin-top:8px}" +
    "#pwa-pop .x{position:absolute;top:6px;right:8px;background:none;border:0;color:#e6edfb;font-size:20px;cursor:pointer;opacity:.7}" +
    "#pwa-pop .go{margin-top:12px;width:100%;padding:11px 0;font-size:16px;font-weight:700;border:0;border-radius:10px;" +
    "background:#4f7df0;color:#fff;cursor:pointer}" +
    "#pwa-pop .bar{position:absolute;left:0;bottom:0;height:3px;background:#4f7df0;width:100%}";
  D.head.appendChild(css);

  function close() {
    clearTimeout(timer);
    D.removeEventListener("pointerdown", outside, true);
    if (!box) return;
    var b = box; box = null;
    b.classList.remove("on");
    setTimeout(function () { b.remove(); }, 300);
  }
  function outside(e) { if (box && !box.contains(e.target)) close(); }

  function render(mode, seconds) {
    var T = tx(), icon = D.querySelector('link[rel="apple-touch-icon"]');
    var body;
    if (mode === "help") {
      var k = isIOSChrome ? "ios_chrome" : isIOS ? "ios" : isMacSafari ? "mac" : "android";
      body = "<p><b>" + md(T[k + "_h"]) + "</b><br>" + md(T[k]) + "</p>";
      body = '<div class="help">' + body + "</div>";
    } else {
      body = '<button class="go" type="button">' + md(T.btn) + "</button>";
    }
    if (!box) {
      box = D.createElement("div");
      box.id = "pwa-pop";
      box.setAttribute("role", "dialog");
      D.body.appendChild(box);
      requestAnimationFrame(function () { requestAnimationFrame(function () { box && box.classList.add("on"); }); });
    }
    box.innerHTML =
      '<button class="x" type="button" aria-label="' + md(T.close) + '">×</button>' +
      '<div class="row">' + (icon ? '<img alt="" src="' + icon.href + '">' : "") +
      "<div><h4>" + md(T.title) + "</h4><p>" + md(T.text) + "</p></div></div>" + body +
      '<div class="bar"></div>';
    box.querySelector(".x").onclick = close;
    var go = box.querySelector(".go");
    if (go) go.onclick = function () { W.__pwaInstall(); };
    // laika josla + automātiska aizvēršana
    clearTimeout(timer);
    var bar = box.querySelector(".bar");
    bar.style.transition = "none"; bar.style.width = "100%";
    requestAnimationFrame(function () { requestAnimationFrame(function () {
      bar.style.transition = "width " + seconds + "s linear"; bar.style.width = "0%";
    }); });
    timer = setTimeout(close, seconds * 1000);
    // pieskāriens jebkur ārpus loga to aizver (klausītājs tiek pievienots pēc mirkļa,
    // lai to neaizvērtu tas pats klikšķis, kas logu atvēra)
    D.removeEventListener("pointerdown", outside, true);
    setTimeout(function () { if (box) D.addEventListener("pointerdown", outside, true); }, 50);
  }

  W.__pwaState = function () { return installed ? "installed" : deferred ? "ready" : isIOS ? "ios" : "other"; };
  function fire() {
    var p = deferred; deferred = null;
    close();
    p.prompt();               // atver Chrome sistēmas logu "Instalēt lietotni" (tas pats, kas ⋮ → Instalēt)
    p.userChoice.then(function (r) { if (r && r.outcome === "accepted") installed = true; }).catch(function () {});
  }
  W.__pwaInstall = function () {
    if (installed) return;
    if (deferred) { fire(); return; }
    // Apple (iOS Safari/Chrome, Mac Safari) neļauj lapai pašai instalēties → precīza pamācība
    if (isIOS || isMacSafari) { render("help", 25); return; }
    // Pārlūks vēl nav sagatavojis instalēšanu (piem., lapa tikko atvērta) → pagaida līdz 3 s
    // (pieskāriena atļauja prompt() izsaukšanai ilgst ~5 s), tad parāda norādi.
    var t0 = Date.now();
    (function wait() {
      if (deferred) { fire(); return; }
      if (installed) return;
      if (Date.now() - t0 > 3000) { render("help", 15); return; }
      setTimeout(wait, 100);
    })();
  };

  if (!installed) {
    // neliela pauze, lai pārlūks paspēj nolasīt manifestu un sagatavot instalēšanas logu
    setTimeout(function () { if (!installed) render("offer", 7); }, 600);
  }
})();
"""

PWA_HTML = """
<script>
(function () {
  var P = window.parent;
  // Streamlit Cloud rāda lietotni iekšējā rāmī (iframe). Chrome instalēšanu piedāvā tikai
  // augšējai lapai, tāpēc manifestu un instalēšanas loģiku ievieto augšējā lapā (ja tā ir
  // tajā pašā domēnā un pieejama), citādi – Streamlit lapā.
  var W = P;
  try { if (P.top !== P && P.top.document.head) W = P.top; } catch (e) { W = P; }
  var D = W.document;
  D.documentElement.setAttribute("data-pwa-texts", __TEXTS__);   // atjauno tekstus, ja mainīta valoda
  P.__pwaTarget = W;                                              // sānjoslas poga izsauc W.__pwaInstall
  if (D.getElementById("pwa-manifest")) return;                   // viss pārējais tikai vienreiz
  var base = new URL("app/static/", P.location.origin + P.location.pathname.replace(/[^/]*$/, "")).href;
  // Noņem citus manifestus (piem., hostinga paša), lai Chrome izmanto mūsējo
  Array.prototype.forEach.call(D.querySelectorAll('link[rel="manifest"]'), function (l) { l.remove(); });
  function add(tag, attrs) {
    var el = D.createElement(tag);
    for (var k in attrs) el.setAttribute(k, attrs[k]);
    D.head.appendChild(el);
    return el;
  }
  // instalēšanas loģika vispirms, lai "beforeinstallprompt" netiktu palaists garām
  add("script", {}).textContent = __PWA_JS__;
  add("link", { id: "pwa-manifest", rel: "manifest", href: base + "manifest.json" });
  add("link", { rel: "apple-touch-icon", href: base + "apple-touch-icon.png" });
  add("meta", { name: "theme-color", content: "#0b1d44" });
  add("meta", { name: "mobile-web-app-capable", content: "yes" });
  add("meta", { name: "apple-mobile-web-app-capable", content: "yes" });
  add("meta", { name: "apple-mobile-web-app-status-bar-style", content: "black-translucent" });
  add("meta", { name: "apple-mobile-web-app-title", content: "Līgumu analīze" });
  if ("serviceWorker" in W.navigator) {
    W.navigator.serviceWorker.register(base + "sw.js").catch(function () {});
  }
})();
</script>
"""
_pwa_texts = {k: T.get("pwa_" + k, TR["English"]["pwa_" + k]) for k in
              ("btn", "installed", "close", "ios_h", "ios", "ios_chrome_h", "ios_chrome",
               "mac_h", "mac", "android_h", "android")}
_pwa_texts["title"] = T.get("pwa_pop_title", TR["English"]["pwa_pop_title"])
_pwa_texts["text"] = T.get("pwa_pop_text", TR["English"]["pwa_pop_text"])
components.html(
    PWA_HTML.replace("__TEXTS__", json.dumps(json.dumps(_pwa_texts, ensure_ascii=False), ensure_ascii=False).replace("</", "<\\/"))
            .replace("__PWA_JS__", json.dumps(PWA_JS).replace("</", "<\\/")),
    height=0,
)

# ---------------------------------------------------------------
# Sesijas pārvaldība (st.session_state): analīžu vēsture un "Sākt no jauna"
# Struktūra:
#   history      - saraksts ar analīzēm šajā sesijā (līdz 10): laiks, nosaukums, režīms, rezultāts, līguma dati
#   active       - kuras analīzes rezultāts tiek rādīts
#   result/terms - aktīvās analīzes rezultāts un līguma dati (tos rāda tālāk esošais kods)
#   calc_store   - kalkulatora ievades dati katrai analīzei atsevišķi
# ---------------------------------------------------------------
def session_add(result, terms, label, mode, contract_text=""):
    """Pievieno jaunu analīzi vēsturei un padara to par aktīvo."""
    hist = st.session_state.setdefault("history", [])
    now = datetime.datetime.now()
    entry = {
        "id": now.strftime("%H%M%S%f"), "time": now.strftime("%H:%M:%S"),
        "label": label, "mode": mode, "result": result, "terms": terms or {}, "contract_text": contract_text,
    }
    hist.append(entry)
    del hist[:-10]                                   # glabā pēdējās 10 analīzes
    st.session_state["history"] = hist
    st.session_state["active"] = len(hist) - 1
    st.session_state["history_pick"] = len(hist) - 1
    st.session_state["result"] = entry["result"]
    st.session_state["terms"] = entry["terms"]


def session_select(i):
    """Pārslēdz aktīvo analīzi uz vēstures ierakstu (bez jauna Gemini izsaukuma)."""
    hist = st.session_state.get("history", [])
    if 0 <= i < len(hist):
        st.session_state["active"] = i
        st.session_state["result"] = hist[i]["result"]
        st.session_state["terms"] = hist[i]["terms"]


def current_entry():
    hist = st.session_state.get("history", [])
    i = st.session_state.get("active")
    return hist[i] if isinstance(i, int) and 0 <= i < len(hist) else None


def reset_session():
    """Notīra analīzes, vēsturi, kalkulatoru un ievades laukus. API atslēga un valoda paliek.
    Kešs netiek dzēsts, tāpēc jaunu izmaksu nav."""
    keep = {"api_key_input", "custom_lang"}
    gen = st.session_state.get("uploader_gen", 0) + 1
    for k in list(st.session_state.keys()):
        if k not in keep:
            del st.session_state[k]
    st.session_state["uploader_gen"] = gen           # jauna atslēga = tīrs failu augšupielādes lauks
st.title(T["title"])
st.markdown(T["subtitle"])
# Vieta analīzes rezultātu cilnēm: lapas augšā (virs augšupielādes), aizpilda koda beigās
results_slot = st.container()

# Sānu josla
st.sidebar.header(T["sidebar_header"])
def _saved_api_key():
    """Noklusētā atslēga no Streamlit Secrets vai vides mainīgā GEMINI_API_KEY (nav jāievada katru reizi)."""
    try:
        k = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:                       # secrets.toml nav
        k = ""
    return str(k or os.environ.get("GEMINI_API_KEY", "")).strip()


_SAVED_KEY = _saved_api_key()
api_key = st.sidebar.text_input(T["api_label"], type="password", key="api_key_input",
                                placeholder="✅ saglabāta (var atstāt tukšu)" if _SAVED_KEY else "") or _SAVED_KEY
st.sidebar.button(T["reset_btn"], on_click=reset_session, use_container_width=True)

# PWA instalēšanas poga sānu joslas apakšā (loģika: window.__pwaInstall galvenajā lapā, sk. PWA_JS)
PWA_BTN_HTML = """
<style>
  body { margin: 0; font-family: "Source Sans Pro", sans-serif; }
  button { width: 100%; padding: 9px 0; font-size: 15px; font-weight: 600; border-radius: 8px; cursor: pointer;
           border: 1px solid #4f7df0; background: #4f7df0; color: #fff; }
  button:hover { filter: brightness(1.1); }
  button:disabled { background: transparent; color: #4f7df0; cursor: default; filter: none; }
</style>
<button id="b" type="button"></button>
<script>
(function () {
  var b = document.getElementById("b"), L = __LABELS__;
  function tgt() { var p = window.parent; return p.__pwaTarget || p; }   // augšējā lapa (sk. PWA_HTML)
  function sync() {
    var P = tgt();
    var done = P.__pwaState && P.__pwaState() === "installed";
    b.textContent = done ? L.installed : L.btn;
    b.disabled = done;
  }
  b.addEventListener("click", function () { var P = tgt(); if (P.__pwaInstall) P.__pwaInstall(); setTimeout(sync, 500); });
  sync();
  setInterval(sync, 2000);
})();
</script>
"""
st.sidebar.divider()
with st.sidebar:
    components.html(
        PWA_BTN_HTML.replace("__LABELS__", json.dumps({
            "btn": T.get("pwa_btn", TR["English"]["pwa_btn"]),
            "installed": T.get("pwa_installed", TR["English"]["pwa_installed"]),
        }, ensure_ascii=False)),
        height=44,
    )
uploaded_file = st.file_uploader(
    T["uploader"],
    type=["docx", "pdf", "png", "jpg", "jpeg"],
    key=f"uploader_{st.session_state.get('uploader_gen', 0)}",
)

with st.expander(T["paste_label"], expanded=False):
    # Augstums pēc teksta garuma, lai pilns teksts būtu redzams bez ritināšanas
    _pt = st.session_state.get("pasted_text", "") or ""
    _lines = sum(max(1, -(-len(l) // 90)) for l in _pt.split("\n"))
    pasted_text = st.text_area(
        T["paste_label"], height=min(max(200, 26 * _lines + 40), 5000),
        key="pasted_text", label_visibility="collapsed",
    )

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


# ---------------------------------------------------------------
# Determinisms: temperature 0.0, top_p 1.0, top_k 1 un seed (ja SDK to atbalsta)
# ---------------------------------------------------------------
TEMPERATURE, TOP_P, SEED = 0.0, 1.0, 20240601
try:
    from google.ai import generativelanguage as _gl
    SEED_SUPPORTED = "seed" in _gl.GenerationConfig.pb(_gl.GenerationConfig()).DESCRIPTOR.fields_by_name
except Exception:
    SEED_SUPPORTED = False   # vecāks SDK: seed netiek sūtīts; stabilitāti nodrošina temperature 0, top_k 1 un SQLite kešs
CACHE_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "analysis_cache.db")


def det_config(**extra):
    cfg = {"temperature": TEMPERATURE, "top_p": TOP_P, "top_k": 1}
    if SEED_SUPPORTED:
        cfg["seed"] = SEED
    cfg.update(extra)
    return cfg


def _resp_text(resp):
    try:
        return resp.text or ""
    except ValueError:          # bloķēta vai tukša atbilde
        return ""


def _resp_usage(resp):
    um = getattr(resp, "usage_metadata", None)
    return (getattr(um, "prompt_token_count", 0) or 0, getattr(um, "candidates_token_count", 0) or 0)


PROMPT_VER = "v12-stable-types"   # mainot uzvedni, palieliniet; kritēriju/promptu izmaiņas atjaunojas arī automātiski (ae.PROMPT_FINGERPRINT)


def legal_notes_text(codes=("DTIL", "CL")):
    return "\n".join(f"- {c} {n}. pants: {t}" for (c, n), t in ARTICLE_NOTES.items() if c in codes)


def analyze_contract(_api_key, lang_name, kind, content, nonce=0, prompt_ver=PROMPT_VER, progress=None, store=True):
    """Pilna analīze: SQLite kešs -> 1. slānis (kritēriju grupas, objektīvā klasifikācija Python kodā) -> korektūra.
    Atslēga SQLite kešā = līguma SHA-256 + prompta versija + modelis + atskaites valoda.
    nonce != 0 (poga "Jauna analīze") apiet kešu un pārraksta ierakstu."""
    genai.configure(api_key=_api_key)
    chash = ae.content_hash(content)
    pver = f"{prompt_ver}-{ae.PROMPT_FINGERPRINT}"
    db = ae.AnalysisCache(CACHE_DB)
    if not nonce:
        hit = db.get(chash, pver, MODEL_PRIMARY, lang_name)
        if hit:
            return hit["analysis"], hit["laws_source"] + " | 💾 SQLite (0 API)"

    laws_text, laws_source = load_laws()

    def llm(parts, max_tokens):
        resp = GeminiModel(det_config(response_mime_type="application/json", max_output_tokens=max_tokens)).generate_content(parts)
        return _resp_text(resp), _resp_usage(resp)

    ctx = ae.AnalysisContext(
        kind=kind, content=content, lang_name=lang_name, labels=T,
        llm=llm,
        valid_articles=lambda code: set(article_index(code).keys()),
        legal_notes=legal_notes_text(), legal_notes_noma=legal_notes_text(("CL",)), progress=progress,
    )
    report, info = ae.run_analysis(ctx)
    if not report.strip() or "## " not in report:
        raise ValueError("Gemini atgrieza tukšu atbildi")  # kļūda netiek saglabāta kešatmiņā
    main_text = linkify(report)

    # Neitrālā korektūra (2 caurlaides, JSON; Python numurē)
    try:
        proof_text, proof_ok = ae.run_proofread(ctx, T, start=ctx.neutral_count + 1, with_head=ctx.neutral_count == 0)
    except Exception:
        proof_text, proof_ok = "", False
    analysis_all = main_text + ("\n\n" + proof_text if proof_text.strip() else "")
    cost = cost_usd(ctx.tokens_in, ctx.tokens_out)
    source = (laws_source + " | " + info + f" | 💲≈${cost:.3f}")
    if ctx.complete and proof_ok and store:     # salīdzināšanas režīmā atsevišķas analīzes kešā netiek rakstītas
        db.put(chash, pver, MODEL_PRIMARY, lang_name, {"analysis": analysis_all, "laws_source": source})
    return analysis_all, source

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
        f"\n\n{T['disclaimer']}"
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

            if "Riska faktori" in line_str or "🔴" in line_str:
                color = RGBColor(220, 38, 38)
            elif "Brīdinājumi" in line_str or "🟡" in line_str:
                color = RGBColor(217, 119, 6)
            elif "Ieteikumi" in line_str or "🟢" in line_str:
                color = RGBColor(22, 163, 74)
            elif "Naudas" in line_str or "💶" in line_str:
                color = RGBColor(8, 145, 178)
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

            if "Riska faktori" in raw or "🔴" in raw:
                pdf.set_text_color(220, 38, 38)
            elif "Brīdinājumi" in raw or "🟡" in raw:
                pdf.set_text_color(217, 119, 6)
            elif "Ieteikumi" in raw or "🟢" in raw:
                pdf.set_text_color(22, 163, 74)
            elif "Naudas" in raw or "💶" in raw:
                pdf.set_text_color(8, 145, 178)
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
_lv = T["report_lang_name"] == "latviešu"
_MODES = {
    "one": "1️⃣ " + ("Viena analīze" if _lv else "Single analysis"),
    "three": "🔬 " + ("3 analīzes un salīdzinājums (3× izmaksas)" if _lv else "3 analyses and comparison (3× cost)"),
    "test": T["test_mode_label"],
}
_mode = st.radio(
    "Režīms" if _lv else "Mode", list(_MODES), format_func=_MODES.get, key="run_mode", horizontal=True,
    help=("Trīs analīzes darbojas bez keša un parāda, kuri punkti ir stabili un kuri mainās." if _lv else
          "Three analyses bypass the cache and show which points are stable and which change."),
)
test_mode = _mode == "test"
compare_mode = _mode == "three"
force_new = st.checkbox(T["force_new_label"], key="force_new", disabled=_mode != "one")
COMPARE_RUNS = 3
analyze_button = st.button(T["analyze_btn"], type="primary")


UNSUPPORTED_PDF_SCRIPTS = re.compile(
    "[\u0590-\u08FF\u0900-\u0DFF\u0E00-\u0EFF\u1000-\u109F\u1780-\u17FF\u3000-\u9FFF\uAC00-\uD7AF]"
)


def recs_last(text):
    """Pārvieto sadaļu 🟢 (ieteiktie labojumi) aiz 🔵 apsveramajiem punktiem, bet pirms ⚪ neitrālajām
    piezīmēm, lai Word un PDF atskaitē secība sakristu ar lapu."""
    parts = re.split(r"(?m)^(?=## )", text)
    recs = [p for p in parts if p.startswith("## ") and "🟢" in p.split("\n", 1)[0]]
    if not recs:
        return text
    rest = [p for p in parts if p not in recs]
    pos = len(rest)
    for i, p in enumerate(rest):
        if p.startswith("## ") and any(e in p.split("\n", 1)[0] for e in ("🔵", "🔴", "🟡", "💶", "📊")):
            pos = i + 1
    recs = [r if r.endswith("\n\n") else r.rstrip("\n") + "\n\n" for r in recs]
    rest = [r if (not r.startswith("## ") or r.endswith("\n\n")) else r.rstrip("\n") + "\n\n" for r in rest]
    return "".join(rest[:pos] + recs + rest[pos:]).rstrip() + "\n"


# ---------------------------------------------------------------
# Labotais līgums: ieteiktie labojumi (🟢) tiek ievietoti līguma tekstā un izcelti sarkanā krāsā
# ---------------------------------------------------------------
FIX_RED = (220, 38, 38)
FIX_GRAY = (120, 120, 120)
ADD_WORDS = re.compile(r"pievieno|papildin|add|insert|добав|дополн|включ", re.I)
SECTION_WORDS = re.compile(r"sadaļ|section|раздел|nodaļ|chapter|глав", re.I)
QUOTE_RE = re.compile(r"[„“«\"]([^„“«”»\"]{3,})[”“»\"]")


def parse_fixes(text):
    """No sadaļas 🟢 izvelk labojumus: kods, punkta numurs, apraksts, piedāvātais formulējums."""
    fixes = []
    for head, body in split_sections(text):
        if not head or "🟢" not in head:
            continue
        for raw in body.split("\n"):
            line = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)", "", raw).replace("**", "").strip()
            m = re.match(r"^(?:[-*•]\s*)?([KBA]\d+)\.?\s*(.*)$", line)
            if not m:
                continue
            code, rest = m.group(1), m.group(2)
            quotes = list(QUOTE_RE.finditer(rest))
            if not quotes:
                continue
            q = quotes[-1]
            ref, _, desc = rest[:q.start()].partition(":")
            desc = re.sub(r"[^.:]*:\s*$", "", desc).strip(" .:")      # noņem "Piedāvātais formulējums:"
            num = re.match(r"\s*(\d+(?:\.\d+)*)", ref)
            fixes.append({
                "code": code, "ref": ref.strip(), "num": num.group(1) if num else None,
                "desc": desc, "text": q.group(1).strip(),
                "add": bool(ADD_WORDS.search(desc)) or bool(SECTION_WORDS.search(ref)),
            })
    return fixes


def _num_re(num):
    return re.compile(rf"^\s*{re.escape(num)}(?!\d)(?!\.\d)")


def build_fixed_blocks(contract_text, fixes):
    """Saraksts ar blokiem (veids, teksts, kods): 'orig' - nemainīts, 'del' - aizstāts (pārsvītrots),
    'new' - ievietots labojums (sarkans), 'head' - papildinājumu virsraksts."""
    lines = [l.rstrip() for l in contract_text.split("\n")]
    num_any = re.compile(r"^\s*\d+(?:\.\d+)*\.?\s")
    deleted, after, extra = set(), {}, []
    for f in fixes:
        idx = None
        if f["num"]:
            rx = _num_re(f["num"])
            idx = next((i for i, l in enumerate(lines) if rx.match(l)), None)
        if idx is None:
            extra.append(f)
            continue
        if f["add"]:
            sub, j = re.compile(rf"^\s*{re.escape(f['num'])}\.\d"), idx
            for k in range(idx + 1, len(lines)):
                if not lines[k].strip():
                    continue
                if sub.match(lines[k]) or not num_any.match(lines[k]):
                    j = k
                else:
                    break
            after.setdefault(j, []).append((f["text"], f["code"]))
        else:
            deleted.add(idx)
            new = f["text"]
            if not re.match(r"^\s*\d", new):
                new = f"{f['num']}. {new}"
            after.setdefault(idx, []).append((new, f["code"]))
    blocks = []
    for i, l in enumerate(lines):
        if l.strip():
            blocks.append(("del" if i in deleted else "orig", l.strip(), None))
        for new, code in after.get(i, []):
            blocks.append(("new", new, code))
    if extra:
        blocks.append(("head", T["fx_extra"], None))
        for f in extra:
            blocks.append(("new", (f["ref"] + ": " if f["ref"] else "") + f["text"], f["code"]))
    return blocks


def _is_heading(text):
    letters = [c for c in text if c.isalpha()]
    return len(text) < 90 and letters and sum(c.isupper() for c in letters) / len(letters) > 0.7


def fixed_contract_docx(blocks):
    doc = docx.Document()
    h = doc.add_heading(T["fx_title"], level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"{T['fx_legend']} · {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}")
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(*FIX_GRAY)
    for kind, text, code in blocks:
        if kind == "head":
            hh = doc.add_heading(text, level=2)
            if hh.runs:
                hh.runs[0].font.color.rgb = RGBColor(*FIX_RED)
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        if kind == "del":
            run.font.strike = True
            run.font.color.rgb = RGBColor(*FIX_GRAY)
        elif kind == "new":
            run.font.color.rgb = RGBColor(*FIX_RED)
            tag = p.add_run(f"  [{code}]")
            tag.font.size = Pt(8)
            tag.bold = True
            tag.font.color.rgb = RGBColor(*FIX_RED)
        elif _is_heading(text):
            run.bold = True
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def fixed_contract_pdf(blocks):
    font_path, font_bold_path = find_fonts()
    if not font_path:
        raise FileNotFoundError("DejaVu fonts not found (packages.txt: fonts-dejavu-core).")
    pdf = FPDF()
    pdf.add_font("DejaVu", "", font_path)
    pdf.add_font("DejaVu", "B", font_bold_path)
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()
    pdf.set_font("DejaVu", "B", 14)
    pdf.set_text_color(30, 41, 59)
    pdf.multi_cell(0, 9, T["fx_title"], align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("DejaVu", "", 8.5)
    pdf.set_text_color(*FIX_GRAY)
    pdf.multi_cell(0, 5, f"{T['fx_legend']} · {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}",
                   align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    for kind, text, code in blocks:
        if kind == "head":
            pdf.ln(4)
            pdf.set_font("DejaVu", "B", 11)
            pdf.set_text_color(*FIX_RED)
            pdf.multi_cell(0, 6.5, text, new_x="LMARGIN", new_y="NEXT")
            continue
        if kind == "del":
            pdf.set_font("DejaVu", "S", 10)
            pdf.set_text_color(*FIX_GRAY)
        elif kind == "new":
            pdf.set_font("DejaVu", "", 10)
            pdf.set_text_color(*FIX_RED)
            text = f"{text}  [{code}]"
        else:
            pdf.set_font("DejaVu", "B" if _is_heading(text) else "", 10)
            pdf.set_text_color(30, 30, 30)
        pdf.multi_cell(0, 5.5, text, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1.2)
    return bytes(pdf.output())


def build_downloads(report_text):
    """Izveido Word un PDF failus (baitos), lai tos varētu lejupielādēt vairākkārt."""
    result = {"docx": None, "pdf": None, "docx_err": None, "pdf_err": None}
    report_text = recs_last(report_text)
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
SECTION_CODES = {"🔴": "K", "🟡": "B", "💶": "M", "🔵": "A", "⚪": "N"}


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
    totals = {"🔴": 0, "🟡": 0, "💶": 0, "⚪": 0}
    for head, body in split_sections(text):
        for emoji in totals:
            if head and emoji in head:
                totals[emoji] += count_items(head, body)
    cards = [("🔴", "#dc2626", T["h_crit"]), ("🟡", "#d97706", T["h_warn"]), ("💶", "#0891b2", T["h_money"]), ("⚪", "#6b7280", T["h_neutral"])]
    for col, (emoji, color, label) in zip(st.columns(4), cards):
        name, _, sub = label.partition(" (")                # "Riska faktori (finansiāli un citi zaudējumi)"
        sub = f'<div style="font-size:0.75rem;opacity:0.75">{sub.rstrip(")")}</div>' if sub else ""
        col.markdown(
            f'<div style="border-left:6px solid {color};background:{color}22;border-radius:10px;padding:12px 14px;">'
            f'<div style="font-size:2rem;font-weight:700;color:{color};line-height:1.1">{totals[emoji]}</div>'
            f'<div style="font-size:0.85rem;font-weight:600;opacity:0.95">{name}</div>{sub}</div>',
            unsafe_allow_html=True,
        )


COMPARE_ICON = {"kritisks": "🔴", "bridinajums": "🟡", "nauda": "💶", None: "—"}
COMPARE_STATUS = {
    "stable": ("🟢", "Stabils: visās analīzēs ar to pašu līmeni", "Stable: found in every run at the same level"),
    "level": ("🟠", "Atšķiras līmenis: atrasts visur, bet ar dažādu svarīgumu",
              "Level differs: found in every run, but with different severity"),
    "partial": ("🔴", "Nestabils: atrasts tikai dažās analīzēs", "Unstable: found in only some runs"),
}


def render_compare(cmp):
    """Viena faila vairākas analīzes: kopsavilkums, punktu tabula un ieteikumi uzlabojumiem."""
    lv = T["report_lang_name"] == "latviešu"
    n = cmp["runs"]
    rows = cmp["rows"]
    with st.expander("🔬 " + ("Analīžu salīdzinājums" if lv else "Analysis comparison") + f" ({n}×)", expanded=True):
        st.metric("Stabilitāte" if lv else "Stability", f"{cmp['stability'] * 100:.0f}%",
                  help="Punktu daļa, kas atrasta visās analīzēs ar to pašu līmeni." if lv else
                       "Share of points found in every run at the same level.")
        cnt = ["| | " + " | ".join(f"#{i + 1}" for i in range(n)) + " |", "|---|" + "---|" * n]
        for lvl in ("kritisks", "bridinajums", "nauda"):
            cnt.append(f"| {COMPARE_ICON[lvl]} | " + " | ".join(str(c[lvl]) for c in cmp["counts"]) + " |")
        st.markdown("\n".join(cnt))
        head = ("| Punkts | " if lv else "| Clause | ") + " | ".join(f"#{i + 1}" for i in range(n)) + " | " + ("Statuss" if lv else "Status") + " |"
        lines = [head, "|---|" + "---|" * n + "---|"]
        for r in rows:
            lines.append(f"| {r['ref']} | " + " | ".join(COMPARE_ICON[x] for x in r["levels"]) + f" | {COMPARE_STATUS[r['status']][0]} |")
        st.markdown("\n".join(lines) if rows else "–")
        st.caption(" · ".join(f"{v[0]} {v[1] if lv else v[2]}" for v in COMPARE_STATUS.values()))
        shaky = [r for r in rows if r["status"] != "stable"]
        st.markdown("**" + ("Kā uzlabot" if lv else "How to improve") + "**")
        if not shaky:
            st.success("Visi punkti ir stabili: kritēriji darbojas vienādi katrā analīzē." if lv else
                       "All points are stable: the criteria behave the same in every run.")
        for r in shaky:
            lvls = " / ".join(COMPARE_ICON[x] for x in r["levels"])
            if r["status"] == "partial":
                tip = ("Atrasts " + f"{r['found']}/{n}" + " analīzēs: kritērijs vai līguma formulējums ir neskaidrs. "
                       "Precizē kritērija aprakstu (CHECKLIST) vai līguma punktu; ja punkts ir reāls, tas jāpadara obligāts.") if lv else (
                       f"Found in {r['found']}/{n} runs: the criterion or the contract wording is ambiguous. "
                       "Sharpen the CHECKLIST description or the clause; if the point is real, make it mandatory.")
            else:
                tip = ("Līmenis mainās: jā/nē pazīmes modelis atbild atšķirīgi. Precizē pazīmju definīcijas (classify), lai līmeni nosaka skaidrs noteikums.") if lv else (
                       "Level changes: the model answers the yes/no flags differently. Sharpen the flag definitions (classify) so a clear rule decides the level.")
            with st.expander(f"{COMPARE_STATUS[r['status']][0]} {r['ref']}  {lvls}"):
                st.markdown(tip)
                for i in range(n):
                    if r["levels"][i]:
                        st.markdown(f"**#{i + 1} {COMPARE_ICON[r['levels'][i]]}** {r['problems'][i]}"
                                    + (f"  \n*{r['fixes'][i]}*" if r["fixes"][i] else ""))
                    else:
                        st.markdown(f"**#{i + 1}** — " + ("nav atrasts" if lv else "not found"))
        with st.expander("📄 " + ("Visas atskaites" if lv else "All reports")):
            tabs = st.tabs([f"#{i + 1}" for i in range(n)])
            for t, rep_text in zip(tabs, cmp.get("reports", [])):
                with t:
                    st.markdown(rep_text)


def render_sections(text, only=None, skip=()):
    """Katra sadaļa ir sakļaujama. Pēc noklusējuma atvērtas: kopējais novērtējums, kritiskie, brīdinājumi.
    only: rādīt tikai sadaļas ar šīm emocijzīmēm; skip: izlaist sadaļas ar šīm emocijzīmēm."""
    for head, body in split_sections(text):
        if only is not None and not (head and any(e in head for e in only)):
            continue
        if head and any(e in head for e in skip):
            continue
        if head is None:
            st.markdown(body)
            continue
        n = count_items(head, body)
        name, sep, sub = head.partition(" (")              # skaits pirms paskaidrojuma: "Riska faktori · 3 (…)"
        label = name + (f" · {n}" if n else "") + sep + sub
        with st.expander(label, expanded=any(e in head for e in ("📊", "🔴", "🟡", "🟢"))):
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
    # Civillikums: līgumsods, procenti, noma un īre (īsi pārstāsti pēc likumi.lv teksta)
    ("CL", 1716): "Līgumsods par saistību neizpildi ir konkrēta naudas summa, ko nedrīkst noteikt atkārtotu vai pieaugošu maksājumu veidā. Līgumsods par nepienācīgu izpildi vai kavējumu var būt pieaugošs, bet kopumā ne vairāk par 10 procentiem no pamatparāda.",
    ("CL", 1717): "Līgumsoda apmēru nosaka līdzēji, bet tam jābūt samērīgam un atbilstošam godīgai darījumu praksei.",
    ("CL", 1764): "Procentu apmēru var noteikt līdzēji, bet nesamērīgi un godīgai darījumu praksei neatbilstoši procenti ir prettiesiski.",
    ("CL", 1765): "Procentu apmērs cieši jānosaka aktā vai darījumā; ja tas nav noteikts, piemēro likumiskos procentus.",
    ("CL", 2112): "Noma vai īre ir līgums, ar ko viena puse piešķir otrai lietas lietošanu par zināmu nomas vai īres maksu.",
    ("CL", 2116): "Nomnieks, nododot lietu tālāk trešai personai, nedrīkst to nodot citādai lietošanai vai uz ilgāku laiku, kā uz viņa paša nomas laiku.",
    ("CL", 2117): "Apakšnomas līgums pastāv patstāvīgi; apakšnomnieks drīkst maksāt tieši pirmajam iznomātājam, ciktāl pirmais nomnieks ir parādā.",
    ("CL", 2120): "Nomas maksai jābūt patiesai atlīdzībai par lietas lietošanu, to nedrīkst nolikt tikai izskata pēc.",
    ("CL", 2122): "Nomas un īres maksai jābūt noteiktai.",
    ("CL", 2126): "Ierakstot nomas vai īres līgumu zemesgrāmatā, nomnieks iegūst lietu tiesību, kas ir spēkā arī pret trešām personām.",
    ("CL", 2128): "Abām pusēm jāpiegriež līgumam visa rūpība; tās atbild viena otrai par zaudējumiem, kas noticis ļaunprātības vai neuzmanības dēļ, bet nejauši zaudējumi nav jāatlīdzina.",
    ("CL", 2134): "Iznomātājam jānodod lieta ar piederumiem tādā stāvoklī, lai nomnieks varētu iegūt visu labumu, ko varēja sagaidīt.",
    ("CL", 2138): "Uz lietas gulošās nastas un apgrūtinājumi, ja nav norunāts citādi, jānes iznomātājam.",
    ("CL", 2140): "Nepieciešamie un derīgie izdevumi, ko nomnieks taisījis lietai, iznomātājam jāatlīdzina.",
    ("CL", 2141): "Nomnieks maksā nomas maksu līgumā noteiktos termiņos; kavējot jāmaksā nokavējuma procenti.",
    ("CL", 2144): "Ja nomnieks patvaļīgi atsakās no nomas pirms termiņa, no viņa var prasīt visas nomas maksas samaksu; ja izbeigšanai ir likumīgs pamats, maksā tikai par patieso nomas laiku.",
    ("CL", 2151): "Nomnieks drīkst lietot priekšmetu tikai atbilstoši mērķim, kas bija paredzēts iznomājot.",
    ("CL", 2152): "Pēc nomas beigām nomniekam jāatdod lieta ar piederumiem, cik iespējams labā stāvoklī.",
    ("CL", 2155): "Ja nomniekam bijuši likumīgi izdevumi lietas labā, viņš var lietu aizturēt līdz to atlīdzināšanai.",
    ("CL", 2156): "Ja nomniekam ar līgumu uzlikts izdarīt uzlabojumus, no viņa var prasīt to izpildi jau pirms nomas beigām.",
    ("CL", 2165): "Nomas līgums uz noteiktu termiņu izbeidzas, kad termiņš notecējis.",
    ("CL", 2166): "Uz nenoteiktu laiku noslēgts nekustama īpašuma nomas līgums, ja nav norunāts citādi, izbeidzas tikai pēc iepriekšēja sešu mēnešu uzteikuma.",
    ("CL", 2170): "Katrs līdzējs var vienpusēji atkāpties no līguma, ja radies pārmērīgs zaudējums.",
    ("CL", 2171): "Iznomātājs var prasīt līguma atcelšanu, ja maksa nav samaksāta termiņā, ja viņam rodas neparedzēta nepieciešamība pašam lietot lietu, ja nomnieks bojā lietu vai nomā netiek ievērots līgums.",
    ("CL", 2172): "Nomnieks var prasīt līguma atcelšanu, ja iznomātājs novilcina nodošanu, neizdara vajadzīgos labojumus vai lietai ir trūkumi, kas būtiski kavē lietošanu.",
    ("CL", 2173): "Atkāpjoties no līguma, jādod otrai pusei vajadzīgais laiks. Iznomātājs nedrīkst nomnieku izlikt patvarīgi, pat ja līgumā tāda tiesība pielīgta.",
    ("CL", 2174): "Ja iznomātājs lietu atsavina, jaunajam īpašniekam nomas līgums jāievēro tikai tad, ja tas ierakstīts zemesgrāmatā.",
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
# ---------------------------------------------------------------
# Vēstures izvēlne un kalkulators ar sesijā saglabātiem datiem
# ---------------------------------------------------------------
def render_history_picker():
    """Ja šajā sesijā ir vairākas analīzes, ļauj pārslēgties starp tām (bez jauna Gemini izsaukuma)."""
    hist = st.session_state.get("history", [])
    if len(hist) < 2:
        return

    def _pick():
        session_select(st.session_state["history_pick"])

    st.selectbox(
        "📜 Sesijas vēsture (iepriekšējās analīzes šajā sesijā)",
        options=list(range(len(hist))),
        format_func=lambda i: f"{hist[i]['time']} · {'🧪' if hist[i]['mode'] == 'test' else '📄'} {hist[i]['label']}",
        key="history_pick",
        on_change=_pick,
    )


def render_calculator():
    """Soda naudas kalkulators. Ievades dati tiek saglabāti sesijā katrai analīzei atsevišķi."""
    entry = current_entry()
    terms = (entry["terms"] if entry else st.session_state.get("terms")) or {}
    sig = entry["id"] if entry else "none"
    store = st.session_state.setdefault("calc_store", {})
    saved = store.get(sig, {})
    with st.expander("🧮 Soda naudas un procentu kalkulators", expanded=False):
        if terms:
            st.caption("Lauki aizpildīti automātiski no līguma teksta (ja tas tika atrasts). Pārbaudiet un, ja vajag, izlabojiet.")
        kind = st.radio(
            "Sankcijas veids",
            ["Līgumsods / nokavējuma nauda (% dienā)", "Nokavējuma procenti (% gadā)"],
            index=0 if saved.get("daily", True) else 1,
            key=f"calc_kind_{sig}",
        )
        daily = kind.startswith("Līgumsods")
        c1, c2, c3 = st.columns(3)
        principal = c1.number_input(
            "Parāda summa / īres maksa (EUR)", min_value=0.0,
            value=float(saved.get("principal", terms.get("rent") or 500.0)),
            step=10.0, format="%.2f", key=f"calc_p_{sig}",
        )
        days = c2.number_input(
            "Kavēto dienu skaits", min_value=0, value=int(saved.get("days", 30)), step=1, key=f"calc_d_{sig}",
        )
        rate = c3.number_input(
            "Likme līgumā (% dienā)" if daily else "Likme līgumā (% gadā)", min_value=0.0,
            value=float(saved.get("rate_day", terms.get("rate_day") or 0.1)) if daily else float(saved.get("rate_year", 6.0)),
            step=0.01, format="%.3f", key=f"calc_r_{sig}_{int(daily)}",
        )
        cap_declared = True
        if daily:
            cap_declared = st.checkbox(
                "Līgumā ir noteikts kopējais līgumsoda griesti (ne vairāk kā 10% no parāda)",
                value=bool(saved.get("cap", terms.get("cap"))), key=f"calc_cap_{sig}",
            )

        # Saglabā ievades datus sesijā (nepazūd, pārslēdzot analīzes vēsturē)
        new = dict(saved)
        new.update({"principal": principal, "days": days, "daily": daily})
        new["rate_day" if daily else "rate_year"] = rate
        if daily:
            new["cap"] = cap_declared
        store[sig] = new

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
# ---------------------------------------------------------------
# Čats ar līgumu (testa režīmā bez Gemini; reālajā režīmā ar līguma kontekstu)
# Ziņas glabājas st.session_state["messages"]; katra ziņa pieder konkrētai analīzei (entry)
# ---------------------------------------------------------------
import unicodedata

CHAT_SYSTEM = (
    "Tu esi Latvijas īres tiesību palīgs. Atbildi uz lietotāja jautājumu par zemāk doto līguma tekstu un tā analīzi. "
    "Noteikumi: atbildi īsi un vienkārši latviešu valodā (ja lietotājs raksta citā valodā, atbildi tajā); "
    "balsties tikai uz līguma tekstu un analīzi; norādi līguma punkta numuru; ja līgumā atbildes nav, to pasaki un neizdomā; "
    "tas ir informatīvs skaidrojums, nevis juridiska konsultācija. "
    "Atsauces uz likumiem raksti tikai formātā [[DTIL 12]] vai [[CL 2174]] un tikai tad, ja esi pārliecināts par panta numuru."
)

# (atslēgvārdu saknes bez garumzīmēm, atbilde). Atbilde ar [[DTIL 27]] atsaucēm kļūst par saitēm.
MOCK_ANSWERS = [
    (("majdzivniek", "dzivniek", "suns", "suni", "kakis", "kaki", "kaku", "kakus"),
     "🧪 **Mājdzīvnieki:** šajā līgumā (6. sadaļa) nav atrunāti noteikumi par mājdzīvnieku turēšanu. Tāpēc nav skaidrs, vai suni vai kaķi "
     "drīkst turēt, kas atbild par bojājumiem un kā tiek risināti trokšņa vai tīrības strīdi. Ieteikums: pievienot punktu, piemēram, "
     "„Īrnieks drīkst turēt vienu mājdzīvnieku ar izīrētāja rakstveida piekrišanu un atbild par tā radītajiem bojājumiem.” "
     "Vispārīgie lietošanas noteikumi: [[DTIL 17]]."),
    (("uzteik", "izbeig", "atkap", "beigt ligumu", "iziet no"),
     "🧪 **Uzteikums:** īrniekam ir tiesības vienpusēji atkāpties no līguma, rakstveidā brīdinot izīrētāju vismaz 1 mēnesi iepriekš "
     "(līgumiem līdz 10 gadiem) [[DTIL 27]] [[DTIL 20]]. Izīrētājs var izbeigt līgumu tikai likumā noteiktajos gadījumos [[DTIL 21]], "
     "piemēram, ja īre kavēta vairāk nekā 2 mēnešus [[DTIL 24]]. Pārbaudiet, vai līguma uzteikuma punkts atbilst šiem termiņiem."),
    (("depozit", "drosibas nauda", "drosibas naud"),
     "🧪 **Drošības nauda:** tā nedrīkst pārsniegt divu mēnešu īres maksu un jāatdod ne vēlāk kā dienā, kad telpa tiek atbrīvota, "
     "ja līgumā nav noteikts citādi [[DTIL 12]]. Līguma 4.3. punkts ļauj ieturēt naudu bez saskaņošanas, tāpēc ieteicams prasīt "
     "abpusēji parakstītu defektu aktu."),
    (("zemesgramat", "pardod", "pardosan", "ipasnieks mainas", "jauns ipasnieks"),
     "🧪 **Zemesgrāmata:** ja telpu pārdod, jaunajam īpašniekam līgums ir saistošs tikai tad, ja īres tiesības ir ierakstītas "
     "zemesgrāmatā [[DTIL 28]] [[DTIL 35]]. Pretējā gadījumā viņš var prasīt telpas atbrīvot 2 mēnešu laikā."),
    (("ligumsod", "soda nauda", "nokavej", "kavej", "procent"),
     "🧪 **Līgumsods un procenti:** līgumsods par kavējumu kopumā nevar pārsniegt 10% no parāda [[CL 1716]], un tam jābūt samērīgam "
     "[[CL 1717]]. Konkrētu summu aprēķiniet kalkulatorā lapas apakšā."),
    (("ires maks", "ires maksa", "paaugstin", "palielin", "cena"),
     "🧪 **Īres maksa:** līguma laikā to drīkst mainīt tikai tad, ja līgumā ir paredzēti maksas grozīšanas nosacījumi un kārtība [[DTIL 10]]."),
    (("remont", "bojajum", "salauz"),
     "🧪 **Remonti:** līguma 6.1.2. punktā „sīkie remontdarbi” nav definēti, tāpēc var rasties strīds, kurš maksā par sadzīves tehnikas "
     "vai santehnikas bojājumiem. Izīrētājs savukārt nodrošina telpas lietošanas iespēju visu līguma laiku [[DTIL 18]]."),
]


def _norm(text):
    """Mazie burti bez garumzīmēm, lai atslēgvārdi darbotos arī bez diakritikas."""
    t = unicodedata.normalize("NFKD", text.lower())
    return "".join(c for c in t if not unicodedata.combining(c))


def mock_chat_answer(question):
    """Testa režīma atbilde pēc atslēgvārdiem (bez Gemini, 0.00$)."""
    q = _norm(question)
    for keys, answer in MOCK_ANSWERS:
        if any(k in q for k in keys):
            return answer
    return (
        f"🧪 **Testa atbilde:** jautājums saņemts („{question.strip()}”). Testa režīmā Gemini netiek izsaukts. "
        "Īstajā režīmā atbildi sagatavotu Gemini, balstoties uz līguma tekstu. "
        "Pamēģiniet jautāt par mājdzīvniekiem, uzteikumu, drošības naudu, zemesgrāmatu, līgumsodu, īres maksu vai remontiem."
    )


def chat_answer(question, entry, previous):
    """Reālais režīms: jautājums + līguma konteksts uz Gemini."""
    genai.configure(api_key=api_key)
    model = GeminiModel(genai.GenerationConfig(temperature=0.2, max_output_tokens=800))
    plain = lambda s: re.sub(r"\[([^\]]+)\]\(https?://[^)]+\)", r"\1", s)     # noņem saites, lai nelieki tokeni
    parts = [CHAT_SYSTEM]
    if entry.get("contract_text"):
        parts.append("LĪGUMA TEKSTS:\n" + entry["contract_text"])
    parts.append("IEPRIEKŠ SAGATAVOTĀ ANALĪZE:\n" + plain(entry["result"]["analysis"]))
    if previous:
        parts.append("IEPRIEKŠĒJĀ SARUNA:\n" + "\n".join(
            f"{'Lietotājs' if m['role'] == 'user' else 'Palīgs'}: {plain(m['content'])}" for m in previous[-6:]
        ))
    parts.append("JAUTĀJUMS: " + question)
    return linkify(model.generate_content("\n\n".join(parts)).text)


def render_chat():
    """Čata saskarne zem analīzes rezultātiem."""
    entry = current_entry()
    if not entry:
        return
    msgs = st.session_state.setdefault("messages", [])
    mine = [m for m in msgs if m["entry"] == entry["id"]]
    use_mock = test_mode or entry["mode"] == "test"        # testa analīzei nav īsta līguma, tāpēc vienmēr paraugs

    st.markdown("### 💬 Čats ar līgumu")
    if use_mock:
        st.caption("🧪 Testa režīms: atbildes ir iebūvēti paraugi, Gemini netiek izsaukts (0.00$).")
    else:
        st.caption("Jautājums tiek sūtīts Gemini kopā ar līguma tekstu un analīzi (neliela API maksa par katru jautājumu).")

    for m in mine:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    prompt = st.chat_input("Uzdodiet jautājumu par šo līgumu (piem., par mājdzīvniekiem vai uzteikumu)")
    if prompt:
        msgs.append({"role": "user", "content": prompt, "entry": entry["id"], "mock": use_mock})
        with st.chat_message("user"):
            st.markdown(prompt)
        with st.chat_message("assistant"):
            if use_mock:
                answer = linkify(mock_chat_answer(prompt))
            elif not api_key:
                answer = "⚠️ Ievadiet Gemini API atslēgu sānu joslā vai ieslēdziet testa režīmu."
            else:
                try:
                    with st.spinner("💬 Gemini sagatavo atbildi..."):
                        answer = chat_answer(prompt, entry, mine)
                except Exception as e:
                    answer = f"❌ Neizdevās sagatavot atbildi: {e}"
            st.markdown(answer)
        msgs.append({"role": "assistant", "content": answer, "entry": entry["id"], "mock": use_mock})
        mine = [m for m in msgs if m["entry"] == entry["id"]]

    if mine:   # skaidrojumi uz pantu atsaucēm arī čatā (ilgā turēšana / peles uzvešana)
        inject_law_tips(entry["result"]["analysis"] + "\n" + "\n".join(m["content"] for m in mine))
# ---------------------------------------------------------------
# Peldošs čata logs labajā pusē: 💬 poga, sveiciena burbulis, sarkanais indikators, paslēpšana
# Ziņas glabājas st.session_state["messages"] (tāpat kā iepriekš). Prasa Streamlit 1.39+ (st.container(key=...))
# ---------------------------------------------------------------
import inspect

CHAT_GREETING = "Sveiks. Kādi tev radās jautājumi par šo līgumu?"

CHAT_CSS = """
<style>
@keyframes chatPop { 0% {opacity: 0; transform: translateY(16px) scale(.9);} 60% {opacity: 1; transform: translateY(-4px) scale(1.03);} 100% {opacity: 1; transform: none;} }
@keyframes chatFade { to {opacity: 0; visibility: hidden;} }
@keyframes chatPulse { 0% {box-shadow: 0 0 0 0 rgba(239,68,68,.55);} 70% {box-shadow: 0 0 0 12px rgba(239,68,68,0);} 100% {box-shadow: 0 0 0 0 rgba(239,68,68,0);} }

.st-key-chat_widget { position: fixed !important; right: 18px; bottom: 72px; width: min(390px, calc(100vw - 24px)); z-index: 1000;
  padding: 10px 12px; border-radius: 16px; border: 1px solid rgba(128,128,128,.4); box-shadow: 0 12px 40px rgba(0,0,0,.4);
  background: #ffffff; animation: chatPop .35s ease-out; }

.st-key-chat_launcher, .st-key-chat_launcher_unread { position: fixed !important; right: 18px; bottom: 72px; z-index: 1000; width: 60px; height: 60px; }
.st-key-chat_launcher button, .st-key-chat_launcher_unread button { width: 60px; height: 60px; min-height: 60px; border-radius: 50%; padding: 0;
  background: #2563eb; border: none; box-shadow: 0 6px 20px rgba(37,99,235,.5); }
.st-key-chat_launcher button p, .st-key-chat_launcher_unread button p { color: #ffffff !important; font-size: 26px; line-height: 1; }
.st-key-chat_launcher_unread::after { content: "1"; position: absolute; top: -4px; right: -4px; min-width: 22px; height: 22px; border-radius: 11px;
  background: #ef4444; color: #ffffff; font: 700 12px/22px sans-serif; text-align: center; border: 2px solid #ffffff; pointer-events: none;
  animation: chatPulse 1.8s infinite; }

.st-key-chat_greeting { position: fixed !important; right: 18px; bottom: 144px; z-index: 1001; width: min(280px, calc(100vw - 40px));
  animation: chatPop .5s ease-out, chatFade .6s ease-in 6s forwards; }
.st-key-chat_greeting button { width: 100%; background: #ffffff; border: 1px solid #d1d5db; border-radius: 16px 16px 4px 16px; padding: 10px 14px;
  text-align: left; box-shadow: 0 8px 28px rgba(0,0,0,.3); }
.st-key-chat_greeting button p { color: #111827 !important; font-size: 14px; line-height: 1.4; }

@media (max-width: 480px) { .st-key-chat_widget { right: 3vw; width: 94vw; } }
</style>
"""

CHAT_JS = """
<script>
/*__TOKEN__*/
(function () {
  var P = window.parent, D = P.document;
  function sync() {                      // čata loga krāsas = lapas krāsas (arī pēc tēmas pārslēgšanas)
    try {
      var w = D.querySelector(".st-key-chat_widget");
      if (!w) return;
      var cs = P.getComputedStyle(D.querySelector(".stApp"));
      var bg = cs.backgroundColor;
      if (!bg || bg === "transparent" || bg === "rgba(0, 0, 0, 0)") bg = P.getComputedStyle(D.body).backgroundColor;
      if (!bg || bg === "transparent" || bg === "rgba(0, 0, 0, 0)") bg = "#ffffff";
      w.style.backgroundColor = bg;
      w.style.color = cs.color;
    } catch (e) {}
  }
  sync();
  if (!P.__chatSync) { P.__chatSync = setInterval(sync, 700); }
  setTimeout(function () {               // ritina ziņas uz leju
    try {
      var c = D.querySelector(".st-key-chat_msgs");
      if (!c) return;
      var els = Array.prototype.slice.call(c.querySelectorAll("*"));
      els.push(c);
      els.forEach(function (el) {
        var oy = P.getComputedStyle(el).overflowY;
        if ((oy === "auto" || oy === "scroll") && el.scrollHeight > el.clientHeight + 4) { el.scrollTop = el.scrollHeight; }
      });
    } catch (e) {}
  }, 200);
})();
</script>
"""


def chat_toggle(open_):
    """Atver vai paslēpj čata logu. Atverot, analīzes sveiciens tiek atzīmēts kā redzēts (sarkanais indikators pazūd)."""
    st.session_state["chat_open"] = open_
    e = current_entry()
    if open_ and e:
        seen = st.session_state.setdefault("chat_seen", [])
        if e["id"] not in seen:
            seen.append(e["id"])


def render_chat_widget():
    """Peldošs čata logs labajā apakšējā stūrī (saruna saglabājas arī paslēpjot)."""
    if "key" not in inspect.signature(st.container).parameters:
        render_chat()                       # vecāka Streamlit versija: parastais čats lapas apakšā
        return
    entry = current_entry()
    if not entry:
        return
    st.markdown(CHAT_CSS, unsafe_allow_html=True)
    msgs = st.session_state.setdefault("messages", [])
    if not any(m["entry"] == entry["id"] for m in msgs):
        msgs.append({"role": "assistant", "content": CHAT_GREETING, "entry": entry["id"], "mock": False, "greeting": True})
    mine = [m for m in msgs if m["entry"] == entry["id"]]
    use_mock = test_mode or entry["mode"] == "test"
    is_open = st.session_state.get("chat_open", False)
    unread = entry["id"] not in st.session_state.get("chat_seen", [])

    if not is_open:
        greeted = st.session_state.setdefault("chat_greeted", [])
        if unread and entry["id"] not in greeted:           # sveiciena burbulis tiek parādīts vienu reizi katrai analīzei
            greeted.append(entry["id"])
            with st.container(key="chat_greeting"):
                st.button("👋 " + CHAT_GREETING, key="chat_greet_btn", on_click=chat_toggle, args=(True,))
        with st.container(key="chat_launcher_unread" if unread else "chat_launcher"):
            st.button("💬", key="chat_open_btn", on_click=chat_toggle, args=(True,))
        return

    with st.container(key="chat_widget"):
        h1, h2 = st.columns([5, 1])
        h1.markdown("**💬 Čats ar līgumu**")
        h2.button("✕", key="chat_close_btn", on_click=chat_toggle, args=(False,))
        if use_mock:
            st.caption("🧪 Testa režīms: iebūvētas atbildes, Gemini netiek izsaukts (0.00$).")
        else:
            st.caption("Jautājums tiek sūtīts Gemini kopā ar līguma tekstu un analīzi.")
        box = st.container(height=280, key="chat_msgs")
        prompt = st.chat_input("Uzdodiet jautājumu par līgumu...", key="chat_widget_input")

    with box:                                                # ziņas tiek zīmētas virs ievades lauka
        for m in mine:
            with st.chat_message(m["role"]):
                st.markdown(m["content"])
        if prompt:
            msgs.append({"role": "user", "content": prompt, "entry": entry["id"], "mock": use_mock})
            with st.chat_message("user"):
                st.markdown(prompt)
            with st.chat_message("assistant"):
                if use_mock:
                    answer = linkify(mock_chat_answer(prompt))
                elif not api_key:
                    answer = "⚠️ Ievadiet Gemini API atslēgu sānu joslā vai ieslēdziet testa režīmu."
                else:
                    try:
                        with st.spinner("💬 Gemini sagatavo atbildi..."):
                            answer = chat_answer(prompt, entry, mine)
                    except Exception as e:
                        answer = f"❌ Neizdevās sagatavot atbildi: {e}"
                st.markdown(answer)
            msgs.append({"role": "assistant", "content": answer, "entry": entry["id"], "mock": use_mock})
            mine = [m for m in msgs if m["entry"] == entry["id"]]

    components.html(CHAT_JS.replace("__TOKEN__", f"{len(mine)}-{int(is_open)}"), height=0)
    inject_law_tips(entry["result"]["analysis"] + "\n" + "\n".join(m["content"] for m in mine))
# ---------------------------------------------------------------
# Kalkulators izvēlētajā saskarnes valodā (pārraksta iepriekšējās penalty_calc un render_calculator definīcijas)
# ---------------------------------------------------------------
def _dec_comma():
    """Decimālkomats latviešu un krievu valodā; citās valodās (arī Gemini tulkotajās) decimālpunkts."""
    return T.get("report_lang_name") in ("latviešu", "krievu")


def _n(x, d=2):
    s = f"{x:,.{d}f}"
    return s.replace(",", " ").replace(".", ",") if _dec_comma() else s


def eur(x):
    return _n(x, 2) + " EUR"


def pct(x):
    s = f"{x:g}"
    return s.replace(".", ",") if _dec_comma() else s


def _label(form, n):
    try:
        return T[form].format(n=n)
    except Exception:
        return TR["English"][form].format(n=n)


def _cl_links():
    """Saites uz Civillikuma pantiem ar nosaukumu izvēlētajā valodā (veidnēs: {cl1716_gen}, {cl1765} u.c.)."""
    def link(n, form="cite_cl", shown=None):
        return f"[{_label(form, shown or n)}]({LAW_LINKS['CL']}#p{n})"
    return {
        "cl1716_gen": link(1716, "cite_cl_gen"), "cl1722_acc": link(1722, "cite_cl_acc"),
        "cl1716": link(1716), "cl1717": link(1717), "cl1764": link(1764), "cl1765": link(1765),
        "cl1724": link(1724, shown="1724¹"),
    }


def tfmt(key, **kw):
    """Aizpilda tulkoto veidni. Ja tulkojumā vietturi sabojāti, izmanto angļu veidni (nekad nenoved pie kļūdas)."""
    vals = {**_cl_links(), **kw}
    for tpl in (T.get(key), TR["English"][key]):
        try:
            return tpl.format(**vals)
        except Exception:
            continue
    return TR["English"][key]


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
        res["days_to_cap"] = 10 / rate
        res["lawful"] = min(total, cap)
        res["excess"] = max(0.0, total - cap)
        res["claimable"] = max(0.0, res["lawful"] - statutory)      # CL 1722.
        if total > cap + 0.005:
            res["level"], res["title"] = "error", T["calc_t_error"]
            res["messages"] = [tfmt("calc_msg_exceed", total=eur(total), cap=eur(cap), excess=eur(res["excess"]))]
            return res
        msgs = []
        if not cap_declared:
            msgs.append(tfmt("calc_msg_nocap", rate=pct(rate), days_to_cap=_n(res["days_to_cap"], 0)))
        if rate > 0.1:
            msgs.append(tfmt("calc_msg_highrate", rate=pct(rate), annual=_n(annual, 0), multiple=_n(multiple, 1)))
        if msgs:
            res["level"], res["title"], res["messages"] = "warning", T["calc_t_warn"], msgs
        else:
            res["level"], res["title"] = "success", T["calc_t_ok"]
            res["messages"] = [tfmt("calc_msg_ok_daily", total=eur(total), cap=eur(cap))]
        return res

    # nokavējuma procenti (% gadā)
    if rate <= 6:
        res["level"], res["title"] = "success", T["calc_t_ok"]
        res["messages"] = [tfmt("calc_msg_ok_annual", rate=pct(rate))]
    elif rate <= 36.5:
        res["level"], res["title"] = "warning", T["calc_t_warn"]
        res["messages"] = [tfmt("calc_msg_warn_annual", rate=pct(rate), multiple=_n(multiple, 1))]
    else:
        res["level"], res["title"] = "error", T["calc_t_risk"]
        res["messages"] = [tfmt("calc_msg_risk_annual", rate=pct(rate), rate_day=pct(round(rate / 365, 2)), multiple=_n(multiple, 1))]
    return res


def render_calculator(boxed=True):
    """Soda naudas kalkulators izvēlētajā valodā. Ievades dati tiek saglabāti sesijā katrai analīzei atsevišķi.
    boxed=False: bez sakļaujamā rāmja (izmanto cilnē)."""
    entry = current_entry()
    terms = (entry["terms"] if entry else st.session_state.get("terms")) or {}
    sig = entry["id"] if entry else "none"
    store = st.session_state.setdefault("calc_store", {})
    saved = store.get(sig, {})
    if not boxed:
        st.markdown(f"### {T['calc_title']}")
    with (st.expander(T["calc_title"], expanded=False) if boxed else st.container()):
        if terms:
            st.caption(T["calc_autofill"])
        opts = [T["calc_kind_daily"], T["calc_kind_annual"]]
        kind = st.radio(
            T["calc_kind"], opts, index=0 if saved.get("daily", True) else 1, key=f"calc_kind_{sig}",
        )
        daily = kind == opts[0]
        c1, c2, c3 = st.columns(3)
        principal = c1.number_input(
            T["calc_principal"], min_value=0.0, value=float(saved.get("principal", terms.get("rent") or 500.0)),
            step=10.0, format="%.2f", key=f"calc_p_{sig}",
        )
        days = c2.number_input(
            T["calc_days"], min_value=0, value=int(saved.get("days", 30)), step=1, key=f"calc_d_{sig}",
        )
        rate = c3.number_input(
            T["calc_rate_day"] if daily else T["calc_rate_year"], min_value=0.0,
            value=float(saved.get("rate_day", terms.get("rate_day") or 0.1)) if daily else float(saved.get("rate_year", 6.0)),
            step=0.01, format="%.3f", key=f"calc_r_{sig}_{int(daily)}",
        )
        cap_declared = True
        if daily:
            cap_declared = st.checkbox(
                T["calc_cap"], value=bool(saved.get("cap", terms.get("cap"))), key=f"calc_cap_{sig}",
            )

        # Saglabā ievades datus sesijā (nepazūd, pārslēdzot analīzes vēsturē vai valodu)
        new = dict(saved)
        new.update({"principal": principal, "days": days, "daily": daily})
        new["rate_day" if daily else "rate_year"] = rate
        if daily:
            new["cap"] = cap_declared
        store[sig] = new

        if principal <= 0 or days <= 0 or rate <= 0:
            st.info(T["calc_enter"])
            return

        r = penalty_calc(principal, days, rate, daily, cap_declared)
        m1, m2, m3 = st.columns(3)
        m1.metric(T["calc_m_contract"], eur(r["total"]))
        if daily:
            m2.metric(T["calc_m_cap"], eur(r["cap"]))
        else:
            m2.metric(T["calc_m_diff"], eur(max(0.0, r["total"] - r["statutory"])))
        m3.metric(T["calc_m_stat"], eur(r["statutory"]))

        body = "\n\n".join(f"- {m}" for m in r["messages"])
        {"success": st.success, "warning": st.warning, "error": st.error}[r["level"]](f"**{r['title']}**\n\n{body}")

        if daily:
            st.markdown(tfmt("calc_both", statutory=eur(r["statutory"]), claimable=eur(r["claimable"])))
        st.caption(T["calc_note"])
# ---------------------------------------------------------------
# Čats izvēlētajā saskarnes valodā (pārraksta iepriekšējās mock_chat_answer, chat_answer un render_chat_widget definīcijas)
# ---------------------------------------------------------------
MOCK_REFS = {
    "dtil10": "[[DTIL 10]]", "dtil12": "[[DTIL 12]]", "dtil17": "[[DTIL 17]]", "dtil18": "[[DTIL 18]]",
    "dtil20": "[[DTIL 20]]", "dtil21": "[[DTIL 21]]", "dtil24": "[[DTIL 24]]", "dtil27": "[[DTIL 27]]",
    "dtil28": "[[DTIL 28]]", "dtil35": "[[DTIL 35]]", "cl1716": "[[CL 1716]]", "cl1717": "[[CL 1717]]",
}

# (veidnes atslēga, atslēgvārdu saknes latviešu / angļu / krievu valodā; salīdzina bez garumzīmēm)
MOCK_TOPICS = [
    ("mock_pets", ("majdzivniek", "dzivniek", "suns", "suni", "kakis", "kaki", "kaku", "kakus",
                   "pet", "dog", "cat", "animal", "животн", "собак", "кошк", "питомц")),
    ("mock_terminate", ("uzteik", "izbeig", "atkap", "terminat", "cancel", "notice", "расторг", "прекрат", "уведомл", "выехат", "съехат")),
    ("mock_deposit", ("depozit", "drosibas naud", "deposit", "залог", "депозит")),
    ("mock_land", ("zemesgramat", "pardod", "pardosan", "land regist", "sell", "sold", "new owner", "земельн", "продаж", "продад", "продан", "собственник")),
    ("mock_penalty", ("ligumsod", "soda naud", "nokavej", "kavej", "procent", "penalt", "fine", "interest", "overdue", "late fee", "неустой", "штраф", "пени", "просроч", "процент")),
    ("mock_rent", ("ires maks", "paaugstin", "palielin", "cena", "rent", "price", "increase", "арендн", "повыш", "цена")),
    ("mock_repairs", ("remont", "bojajum", "salauz", "repair", "damage", "broken", "повреждени", "поломк")),
]


def mock_chat_answer(question):
    """Testa režīma atbilde pēc atslēgvārdiem izvēlētajā valodā (bez Gemini, 0.00$)."""
    q = _norm(question)
    for key, words in MOCK_TOPICS:
        if any(re.search(r"\b" + re.escape(_norm(w)), q) for w in words):
            return tfmt(key, **MOCK_REFS)
    return tfmt("mock_default", question=question.strip(), **MOCK_REFS)


def _chat_model_and_prompt(question, entry, previous):
    """Čata modelis un uzaicinājums (kopīgs parastajai un straumētajai atbildei; saturs nemainās)."""
    genai.configure(api_key=api_key)
    model = GeminiModel(genai.GenerationConfig(temperature=0.2, max_output_tokens=800))
    plain = lambda s: re.sub(r"\[([^\]]+)\]\(https?://[^)]+\)", r"\1", s)     # noņem saites, lai nelieki tokeni
    system = CHAT_SYSTEM.replace(
        "latviešu valodā (ja lietotājs raksta citā valodā, atbildi tajā)", T["report_lang_name"] + " valodā"
    )
    parts = [system]
    if entry.get("contract_text"):
        parts.append("LĪGUMA TEKSTS:\n" + entry["contract_text"])
    parts.append("IEPRIEKŠ SAGATAVOTĀ ANALĪZE:\n" + plain(entry["result"]["analysis"]))
    if previous:
        parts.append("IEPRIEKŠĒJĀ SARUNA:\n" + "\n".join(
            f"{'Lietotājs' if m['role'] == 'user' else 'Palīgs'}: {plain(m['content'])}" for m in previous[-6:]
        ))
    parts.append("JAUTĀJUMS: " + question)
    return model, "\n\n".join(parts)


def chat_answer(question, entry, previous):
    """Reālais režīms: jautājums + līguma konteksts uz Gemini. Atbilde izvēlētajā saskarnes valodā."""
    model, prompt = _chat_model_and_prompt(question, entry, previous)
    return linkify(model.generate_content(prompt).text)


def chat_answer_stream(question, entry, previous):
    """Tas pats pieprasījums kā chat_answer, bet atbilde parādās pa daļām (straumējot). Analīzi neskar."""
    model, prompt = _chat_model_and_prompt(question, entry, previous)
    def chunks():
        for part in model.generate_content(prompt, stream=True):
            try:
                yield part.text
            except ValueError:          # bloķēts vai tukšs fragments
                continue
    yield from linkify_stream(chunks())


def stream_chat_answer(question, entry, previous):
    """Parāda atbildi straumējot un atgriež pilno tekstu (saglabāšanai ziņu vēsturē)."""
    try:
        text = st.write_stream(chat_answer_stream(question, entry, previous))
    except Exception:                   # straumēšana neizdevās: parastā atbilde
        text = chat_answer(question, entry, previous)
        st.markdown(text)
    return text if isinstance(text, str) else "".join(map(str, text))


def render_chat_widget():
    """Peldošs čata logs labajā apakšējā stūrī izvēlētajā saskarnes valodā (saruna saglabājas arī paslēpjot)."""
    if "key" not in inspect.signature(st.container).parameters:
        render_chat()                       # vecāka Streamlit versija: parastais čats lapas apakšā
        return
    entry = current_entry()
    if not entry:
        return
    st.markdown(CHAT_CSS, unsafe_allow_html=True)
    msgs = st.session_state.setdefault("messages", [])
    if not any(m["entry"] == entry["id"] for m in msgs):
        msgs.append({"role": "assistant", "content": "", "entry": entry["id"], "mock": False, "greeting": True})
    mine = [m for m in msgs if m["entry"] == entry["id"]]
    use_mock = test_mode or entry["mode"] == "test"
    is_open = st.session_state.get("chat_open", False)
    unread = entry["id"] not in st.session_state.get("chat_seen", [])

    if not is_open:
        greeted = st.session_state.setdefault("chat_greeted", [])
        if unread and entry["id"] not in greeted:           # sveiciena burbulis tiek parādīts vienu reizi katrai analīzei
            greeted.append(entry["id"])
            with st.container(key="chat_greeting"):
                st.button("👋 " + T["chat_greeting"], key="chat_greet_btn", on_click=chat_toggle, args=(True,))
        with st.container(key="chat_launcher_unread" if unread else "chat_launcher"):
            st.button("💬", key="chat_open_btn", on_click=chat_toggle, args=(True,))
        return

    with st.container(key="chat_widget"):
        h1, h2 = st.columns([5, 1])
        h1.markdown(f"**{T['chat_title']}**")
        h2.button("✕", key="chat_close_btn", on_click=chat_toggle, args=(False,))
        st.caption(T["chat_cap_mock"] if use_mock else T["chat_cap_real"])
        box = st.container(height=280, key="chat_msgs")
        prompt = st.chat_input(T["chat_placeholder"], key="chat_widget_input")

    def shown(m):                                            # sveiciens vienmēr izvēlētajā valodā
        return T["chat_greeting"] if m.get("greeting") else m["content"]

    with box:                                                # ziņas tiek zīmētas virs ievades lauka
        for m in mine:
            with st.chat_message(m["role"]):
                st.markdown(shown(m))
        if prompt:
            msgs.append({"role": "user", "content": prompt, "entry": entry["id"], "mock": use_mock})
            with st.chat_message("user"):
                st.markdown(prompt)
            with st.chat_message("assistant"):
                streamed = False
                if use_mock:
                    answer = linkify(mock_chat_answer(prompt))
                elif not api_key:
                    answer = T["chat_no_key"]
                else:
                    try:
                        answer = stream_chat_answer(prompt, entry, mine)   # tiek parādīta pa daļām
                        streamed = True
                    except Exception as e:
                        answer = tfmt("chat_error", error=e)
                if not streamed:
                    st.markdown(answer)
            msgs.append({"role": "assistant", "content": answer, "entry": entry["id"], "mock": use_mock})
            mine = [m for m in msgs if m["entry"] == entry["id"]]

    components.html(CHAT_JS.replace("__TOKEN__", f"{len(mine)}-{int(is_open)}"), height=0)
    inject_law_tips(entry["result"]["analysis"] + "\n" + "\n".join(shown(m) for m in mine))
def render_chat_tab():
    """Čats ar līgumu kā parasta saskarne cilnē (ziņas tās pašas, kas st.session_state["messages"])."""
    entry = current_entry()
    if not entry:
        return
    msgs = st.session_state.setdefault("messages", [])
    mine = [m for m in msgs if m["entry"] == entry["id"] and not m.get("greeting")]
    use_mock = test_mode or entry["mode"] == "test"
    seen = st.session_state.setdefault("chat_seen", [])
    if entry["id"] not in seen:
        seen.append(entry["id"])

    st.markdown(f"### {T['chat_title']}")
    st.caption(T["chat_cap_mock"] if use_mock else T["chat_cap_real"])
    box = st.container(height=440)
    prompt = st.chat_input(T["chat_placeholder"], key="chat_tab_input")

    with box:                                                # ziņas tiek zīmētas virs ievades lauka
        with st.chat_message("assistant"):
            st.markdown(T["chat_greeting"])
        for m in mine:
            with st.chat_message(m["role"]):
                st.markdown(m["content"])
        if prompt:
            msgs.append({"role": "user", "content": prompt, "entry": entry["id"], "mock": use_mock})
            with st.chat_message("user"):
                st.markdown(prompt)
            with st.chat_message("assistant"):
                streamed = False
                if use_mock:
                    answer = linkify(mock_chat_answer(prompt))
                elif not api_key:
                    answer = T["chat_no_key"]
                else:
                    try:
                        answer = stream_chat_answer(prompt, entry, mine)   # tiek parādīta pa daļām
                        streamed = True
                    except Exception as e:
                        answer = tfmt("chat_error", error=e)
                if not streamed:
                    st.markdown(answer)
            msgs.append({"role": "assistant", "content": answer, "entry": entry["id"], "mock": use_mock})
            mine = [m for m in msgs if m["entry"] == entry["id"] and not m.get("greeting")]

    if mine:   # skaidrojumi uz pantu atsaucēm arī čatā
        inject_law_tips(entry["result"]["analysis"] + "\n" + "\n".join(m["content"] for m in mine))


# ---------------------------------------------------------------
# Čats kā atbīdāms sānu panelis no labās puses (aizstāj peldošo pogu apakšējā labajā stūrī)
# Pārraksta iepriekšējo CHAT_CSS; render_chat_widget() to nolasa izpildes brīdī
# ---------------------------------------------------------------
CHAT_CSS = """
<style>
@keyframes drawerIn { from {transform: translateX(100%);} to {transform: translateX(0);} }
@keyframes tabIn { from {transform: translateX(100%);} to {transform: translateX(0);} }
@keyframes bubbleIn { from {opacity: 0; transform: translateX(12px);} to {opacity: 1; transform: none;} }
@keyframes chatFade { to {opacity: 0; visibility: hidden;} }
@keyframes chatPulse { 0% {box-shadow: 0 0 0 0 rgba(239,68,68,.55);} 70% {box-shadow: 0 0 0 12px rgba(239,68,68,0);} 100% {box-shadow: 0 0 0 0 rgba(239,68,68,0);} }

/* Atvērtais sānu panelis: pilns augstums, izbīdās no labās puses */
.st-key-chat_widget { position: fixed !important; top: 0; right: 0; bottom: 0; height: 100vh; height: 100dvh;
  width: min(420px, 100vw); z-index: 1000002; margin: 0 !important; overflow: hidden;
  padding: calc(16px + env(safe-area-inset-top, 0px)) 16px calc(12px + env(safe-area-inset-bottom, 0px));
  border-left: 1px solid rgba(128,128,128,.35); border-radius: 16px 0 0 16px;
  box-shadow: -12px 0 40px rgba(0,0,0,.30), 0 0 0 100vmax rgba(0,0,0,.12);
  background: #ffffff; animation: drawerIn .3s ease-out; }

/* Ziņu lauks aizpilda paneļa augstumu (ievades lauks paliek apakšā) */
.st-key-chat_msgs { height: calc(100vh - 210px) !important; height: calc(100dvh - 210px) !important; }

/* Aizvērtā stāvoklī: šaura mēlīte labās malas vidū */
.st-key-chat_launcher, .st-key-chat_launcher_unread { position: fixed !important; right: 0; top: calc(50% - 36px);
  z-index: 1000; width: 46px; height: 72px; animation: tabIn .3s ease-out; }
.st-key-chat_launcher button, .st-key-chat_launcher_unread button { width: 46px; height: 72px; min-height: 72px; padding: 0;
  border-radius: 14px 0 0 14px; background: #2563eb; border: none; box-shadow: -4px 4px 16px rgba(37,99,235,.45); }
.st-key-chat_launcher button:hover, .st-key-chat_launcher_unread button:hover { background: #1d4ed8; }
.st-key-chat_launcher button p, .st-key-chat_launcher_unread button p { color: #ffffff !important; font-size: 22px; line-height: 1; }
.st-key-chat_launcher_unread::after { content: "1"; position: absolute; top: -6px; left: -6px; min-width: 22px; height: 22px; border-radius: 11px;
  background: #ef4444; color: #ffffff; font: 700 12px/22px sans-serif; text-align: center; border: 2px solid #ffffff; pointer-events: none;
  animation: chatPulse 1.8s infinite; }

/* Sveiciena burbulis blakus mēlītei */
.st-key-chat_greeting { position: fixed !important; right: 56px; top: calc(50% - 30px); z-index: 1001;
  width: min(280px, calc(100vw - 72px)); animation: bubbleIn .4s ease-out, chatFade .6s ease-in 6s forwards; }
.st-key-chat_greeting button { width: 100%; background: #ffffff; border: 1px solid #d1d5db; border-radius: 16px 4px 16px 16px;
  padding: 10px 14px; text-align: left; box-shadow: 0 8px 28px rgba(0,0,0,.3); }
.st-key-chat_greeting button p { color: #111827 !important; font-size: 14px; line-height: 1.4; }

/* Telefonā panelis aizņem visu ekrāna platumu */
@media (max-width: 480px) { .st-key-chat_widget { width: 100vw; border-radius: 0; border-left: none; } }
</style>
"""
# ---------------------------------------------------------------
# Valodas maiņa pēc analīzes: virsraksti, pantu nosaukumi, avoti un atskaites tiek pārslēgti bez jaunas analīzes (bez Gemini)
# Pārraksta iepriekšējās render_history_picker, render_law_help un law_tip_map definīcijas
# ---------------------------------------------------------------
HEAD_EMOJI_KEYS = {"📊": "h_overall", "🔴": "h_crit", "🟡": "h_warn", "🟢": "h_recs", "🔵": "h_border", "💶": "h_money", "⚪": "h_neutral"}


def localize_analysis(text):
    """Sadaļu virsrakstus un pantu atsauču nosaukumus pārslēdz uz pašreizējo saskarnes valodu."""
    out = []
    for line in text.split("\n"):
        if line.startswith("## "):
            for emoji, key in HEAD_EMOJI_KEYS.items():
                if emoji in line:
                    line = f"## {emoji} {T[key]}"
                    break
        out.append(line)

    def _relabel(m):
        form = "cite_dtil" if m.group(3) == "322216" else "cite_cl"
        return f"[{_label(form, m.group(4))}]({m.group(2)})"

    return LINK_RE.sub(_relabel, "\n".join(out))


def localize_source(s):
    """Likumu avota apraksts pašreizējā valodā."""
    if s.startswith("🧪"):
        return T["test_laws_source"]
    s = re.sub(r"likumi\.lv \(automātiski, (\d+) no (\d+) likumiem\)",
               lambda m: tfmt("src_auto", n=m.group(1), m=m.group(2)), s)
    return (s.replace("rezerves fails likumi.txt", T["src_backup"])
             .replace("nav pieejams", T["src_none"]).replace("neizdevās", T["src_failed"]))


def render_history_picker():
    """Ja šajā sesijā ir vairākas analīzes, ļauj pārslēgties starp tām (bez jauna Gemini izsaukuma)."""
    hist = st.session_state.get("history", [])
    if len(hist) < 2:
        return

    def _pick():
        session_select(st.session_state["history_pick"])

    def _name(h):
        return {"Testa paraugs": T["hist_test"], "Ielīmēts teksts": T["hist_pasted"]}.get(h["label"], h["label"])

    st.selectbox(
        T["hist_label"],
        options=list(range(len(hist))),
        format_func=lambda i: f"{hist[i]['time']} · {'🧪' if hist[i]['mode'] == 'test' else '📄'} {_name(hist[i])}",
        key="history_pick",
        on_change=_pick,
    )


def render_law_help(text):
    """Pieskaramas kartes ar skaidrojumu katram analīzē minētajam pantam (izvēlētajā saskarnes valodā)."""
    seen = []
    for m in LINK_RE.finditer(text):
        key = ("DTIL" if m.group(3) == "322216" else "CL", int(m.group(4)))
        if key not in seen:
            seen.append(key)
    if not seen:
        return
    st.markdown("### " + T["law_head"])
    if T["law_lv_note"]:
        st.caption(T["law_lv_note"])
    cols = st.columns(2)
    for i, (code, num) in enumerate(seen):
        label = _label("cite_dtil" if code == "DTIL" else "cite_cl", num)
        with cols[i % 2]:
            box = st.popover(label, use_container_width=True) if hasattr(st, "popover") else st.expander(label)
            with box:
                note = ARTICLE_NOTES.get((code, num))
                excerpt = article_excerpt(code, num, None)
                if note:
                    st.markdown(f"**{T['law_means']}:** " + note)
                if excerpt:
                    st.markdown(f"**{T['law_official']}:** " + excerpt)
                if not note and not excerpt:
                    st.markdown(T["law_none"])
                st.markdown(f"[{T['law_open']}]({LAW_LINKS[code]}#p{num})")
                st.caption(T["law_disclaimer"])


def law_tip_map(text):
    """Saliek pilnu skaidrojuma tekstu (bez saīsinājumiem) katram analīzē minētajam pantam (ilgā turēšana / uzvešana)."""
    tips = {}
    for m in LINK_RE.finditer(text):
        code = "DTIL" if m.group(3) == "322216" else "CL"
        num = int(m.group(4))
        key = f"{m.group(3)}#p{num}"
        if key in tips:
            continue
        parts = [_label("cite_dtil" if code == "DTIL" else "cite_cl", num)]
        note = ARTICLE_NOTES.get((code, num))
        full = article_excerpt(code, num, None)
        if note:
            parts.append(f"{T['law_means']}: " + note)
        if full:
            parts.append(f"{T['law_official']}: " + full)
        if not note and not full:
            parts.append(T["law_none"])
        tips[key] = "\n\n".join(parts)
    return tips
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


def render_fixed_contract(analysis_text):
    """Izvēle, kurus ieteiktos labojumus ievietot, un labotā līguma lejupielāde (.docx un .pdf)."""
    entry = current_entry() or {}
    eid = entry.get("id", "x")
    st.markdown("---")
    st.markdown(T["fx_header"])
    fixes = parse_fixes(analysis_text)
    if not fixes:
        st.info(T["fx_no_fixes"])
        return
    st.caption(T["fx_intro"])

    keys = [f"fx_{eid}_{f['code']}" for f in fixes]
    c1, c2, _ = st.columns([1, 1, 2])
    if c1.button(T["fx_all"], key=f"fx_all_{eid}", use_container_width=True):
        for k in keys:
            st.session_state[k] = True
    if c2.button(T["fx_clear"], key=f"fx_clear_{eid}", use_container_width=True):
        for k in keys:
            st.session_state[k] = False

    chosen = []
    for f, k in zip(fixes, keys):
        st.session_state.setdefault(k, True)
        label = f"**{f['code']}** · {f['ref'] or '—'}" + (f" — {f['desc']}" if f["desc"] else "")
        if st.checkbox(label, key=k):
            chosen.append(f)
        st.markdown(
            f'<div style="margin:-0.6rem 0 0.5rem 1.8rem;color:#dc2626;font-size:0.9rem">„{_html.escape(f["text"])}”</div>',
            unsafe_allow_html=True,
        )
    st.caption(tfmt("fx_count", n=len(chosen), m=len(fixes)))

    contract = st.text_area(
        T["fx_text_label"], value=entry.get("contract_text", ""), height=180, key=f"fx_text_{eid}",
    )
    if not contract.strip():
        st.warning(T["fx_no_text"])
        return
    if not chosen:
        st.info(T["fx_none_selected"])
        return

    blocks = build_fixed_blocks(contract, chosen)
    today = datetime.date.today()
    col1, col2 = st.columns(2)
    with col1:
        try:
            st.download_button(
                T["fx_dl_word"], data=fixed_contract_docx(blocks),
                file_name=f"Labotais_ligums_{today}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key=f"fx_docx_{eid}",
            )
        except Exception as e:
            st.warning(f"{T['word_fail']}: {e}")
    with col2:
        if UNSUPPORTED_PDF_SCRIPTS.search(contract + " ".join(f["text"] for f in chosen)):
            st.warning(f"{T['pdf_fail']}: font limitation, use the Word file.")
        else:
            try:
                st.download_button(
                    T["fx_dl_pdf"], data=fixed_contract_pdf(blocks),
                    file_name=f"Labotais_ligums_{today}.pdf", mime="application/pdf",
                    use_container_width=True, key=f"fx_pdf_{eid}",
                )
            except Exception as e:
                st.warning(f"{T['pdf_fail']}: {e}")


# ---------------------------------------------------------------
# Analīzes izpilde
# ---------------------------------------------------------------
# ---------------------------------------------------------------
# Testa režīms: iebūvēts parauga rezultāts, Gemini netiek izsaukts (0 €)
# ---------------------------------------------------------------
SAMPLE_CONTRACT = """DZĪVOJAMĀS TELPAS ĪRES LĪGUMS
Rīgā, 2026. gada 1. novembrī

1. PUSES
1.1. Izīrētājs: Jānis Bērziņš, personas kods 010180-12345.
1.2. Īrnieks: Anna Kalniņa, personas kods 020290-54321.

2. LĪGUMA PRIEKŠMETS
2.1. Izīrētājs nodod un Īrnieks pieņem lietošanā divistabu  dzīvokli Brīvības ielā 1-10, Rīgā.
2.2. Līgums noslēgts uz 2 gadiem.

3. ĪRES MAKSA UN MAKSĀJUMI
3.1. Īres maksa ir EUR _____ mēnesī.
3.2. Īres maksu Īrnieks samaksā līdz katra mēneša 5. datumam.
3.4. Īrnieks apmaksā komunālos maksājumus atbilstoši Izīrētāja norādījumiem.

4. DROŠĪBAS NAUDA
4.1. Īrnieks iemaksā drošības naudu viena mēneša īres maksas apmērā.
4.3. Izīrētājs ir tiesīgs ieturēt drošības naudu pēc saviem ieskatiem.
4.5. Drošības naudu atgriež 1 mēneša laikā pēc līguma beigām.

5. LĪGUMA IZBEIGŠANA
5.1. Līgumu var izbeigt 5.7. punktā noteiktajā kārtībā.

6. PUŠU PIENĀKUMI
6.1. Īrnieka pienākumi:
6.1.2. veikt sīkos remontdarbus par saviem līdzekļiem;
6.1.5. ļaut Izīrētājam veikt dzīvokļa apskates.

7. ATBILDĪBA
7.1. Par maksājuma kavējumu Īrnieks maksā 0,5% dienā VAI līgumsodu.

8. NOBEIGUMA NOTEIKUMI
8.1. Līgums sagatavots 2 eksemplāros, pa vienam katrai pusei.
"""


def make_test_analysis():
    """Parauga analīze testēšanai. Atsauces [[DTIL 12]] ar linkify() kļūst par saitēm uz likumi.lv."""
    return f"""## 📊 {T['h_overall']}
Paraugs testēšanai: šo līgumu nevajadzētu parakstīt, kamēr nav labots 2.2. un 4.3. punkts. Galvenie riski ir īres tiesību nenostiprināšana zemesgrāmatā un patvaļīga drošības naudas ieturēšana.

## 🔴 {T['h_crit']}
K1. 2.2. punkts (Zemesgrāmata): Līgumā nav paredzēta īres tiesību nostiprināšana zemesgrāmatā. Ja īpašnieks mainās, jaunais īpašnieks var prasīt telpas atbrīvot 2 mēnešu laikā. [[DTIL 28]] [[DTIL 35]] [[CL 2174]]

K2. 4.3. punkts (Drošības nauda): Izīrētājs drīkst ieturēt naudu bez saskaņošanas ar īrnieku, nepierādot zaudējumus. [[DTIL 12]]

K3. 7.1. punkts (Līgumsods): Līgumsods 0,5% dienā bez kopējiem griestiem ātri pārsniedz 10% no parāda. [[CL 1716]] [[CL 1717]]

## 🟡 {T['h_warn']}
B1. 3.1. punkts (Īres maksa): Īres maksas summa ir tukša (EUR _____). [[DTIL 8]] [[DTIL 10]]

B2. 3.4. punkts (Komunālie maksājumi): Nav noteikta rēķinu uzrādīšanas kārtība. [[DTIL 11]]

B3. 6.1.2. punkts (Sīkie remontdarbi): Nav definēts, kas ir sīkie remontdarbi. [[CL 2162]]

B4. 6.1.5. punkts (Apskates): Nav noteikts minimālais brīdinājuma laiks apskatei. [[DTIL 13]] [[DTIL 18]]

B5. 7.1. punkts (VAI): Vārds VAI starp procentiem un līgumsodu rada neskaidrību. [[CL 1722]]

B6. 4.5. punkts (Depozīta atmaksa): Termiņš 1 mēnesis ir pieļaujams, bet garāks par ierasto praksi; to vērts izrunāt ar otru pusi. [[DTIL 12]]

B7. 1. sadaļa (Bankas konts): Īrniekam nav vietas bankas rekvizītiem, uz kuriem atmaksāt drošības naudu.

## 💶 {T['h_money']}
M1. 3.1. punkts (Īres maksa): Īres maksa mēnesī, maksājama līdz katra mēneša 5. datumam.
M2. 4.5. punkts (Depozīta atmaksa): Drošības naudu atmaksā 1 mēneša laikā pēc telpu nodošanas.

## 🟢 {T['h_recs']}
K1. 2.2. punkts: Pievienot punktu par īres tiesību nostiprināšanu zemesgrāmatā. {T['label_proposed']}: „Puses vienojas nostiprināt īres tiesības zemesgrāmatā. Izīrētājs 10 darba dienu laikā paraksta nostiprinājuma lūgumu.”

K2. 4.3. punkts: Atļaut ieturēt tikai ar abpusēji parakstītu defektu aktu un izdevumu apliecinošiem dokumentiem. {T['label_proposed']}: „Izīrētājs drīkst ieturēt drošības naudu tikai par zaudējumiem, kas fiksēti abpusēji parakstītā aktā un apliecināti ar rēķiniem.”

K3. 7.1. punkts: Noteikt kopējos līgumsoda griestus. {T['label_proposed']}: „Līgumsoda kopsumma nedrīkst pārsniegt 10% no nesamaksātās summas.”

B1. 3.1. punkts: Ierakstīt īres maksas summu cipariem un vārdiem. {T['label_proposed']}: „Īres maksa ir EUR 450,00 (četri simti piecdesmit eiro) mēnesī.”

B2. 3.4. punkts: Noteikt rēķinu kārtību. {T['label_proposed']}: „Izīrētājs līdz katra mēneša 10. datumam nosūta īrniekam rēķinu par komunālajiem maksājumiem ar pievienotām pakalpojumu sniedzēju rēķinu kopijām.”

B3. 6.1.2. punkts: Definēt sīkos remontdarbus. {T['label_proposed']}: „Sīkie remontdarbi ir darbi, kuru izmaksas nepārsniedz EUR 50 vienā reizē (spuldzes, blīves, krānu uzgaļi).”

B4. 6.1.5. punkts: Noteikt brīdinājuma laiku. {T['label_proposed']}: „Izīrētājs par apskati brīdina vismaz 48 stundas iepriekš, izņemot avārijas gadījumus.”

B6. 4.5. punkts: Precizēt depozīta atmaksas termiņu. {T['label_proposed']}: „Drošības naudu izīrētājs atmaksā 10 darba dienu laikā pēc telpu nodošanas.”

B7. 1. sadaļa: Pievienot vietu īrnieka bankas rekvizītiem. {T['label_proposed']}: „Īrnieka bankas konts drošības naudas atmaksai: IBAN _____.”

B5. 7.1. punkts: Atstāt vienu mehānismu: vai nu procentus, vai līgumsodu. {T['label_proposed']}: „Par maksājuma kavējumu īrnieks maksā līgumsodu 0,1% dienā, kopā ne vairāk kā 10% no parāda.”

A1. 4.5. punkts: Saīsināt atmaksas termiņu. {T['label_proposed']}: „Izīrētājs atgriež drošības naudu 10 darba dienu laikā pēc nodošanas-pieņemšanas akta parakstīšanas.”

## ⚪ {T['h_neutral']}
N1. 2.1. punkts: Vārds „istabu” rakstīts ar divām atstarpēm.

N2. 5.1. punkts: Atsauce uz neesošu 5.7. punktu.

N3. 8. sadaļa: Pušu apzīmējumi „Izīrētājs” un „Iznomātājs” lietoti jaukti.
"""


if analyze_button and test_mode:
    sample = linkify(make_test_analysis())
    result = build_downloads(sample + "\n\n## " + T["sources_head"] + "\n" + sources_md())
    result.update({
        "analysis": sample,
        "laws_source": "🧪 TESTA REŽĪMS: iebūvēts paraugs, Gemini netika izsaukts (0.00$)",
        "cache_note": "🧪 Testa režīms: iebūvēts paraugs, Gemini netika izsaukts (0.00$).",
    })
    session_add(result, {"rent": 450.0, "rate_day": 0.5}, "Testa paraugs", "test",
                contract_text=SAMPLE_CONTRACT)
    st.rerun()
if analyze_button:
    # Iepriekšējais rezultāts paliek, līdz tiek saglabāts jauns (nepazūd arī kļūdas gadījumā)
    prev_result = st.session_state.get("result")

    if not api_key:
        st.error(T["err_no_key"])
    elif uploaded_file is None and not pasted_text.strip():
        st.error(T["err_no_file"])
    else:
        try:
            genai.configure(api_key=api_key)
            model = GeminiModel(genai.GenerationConfig(
                    temperature=0,
                    top_p=1,
                    top_k=1
                ))
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
            live = st.empty()                       # starprezultāts parādās, tiklīdz gatavs 1. slānis

            banner = st.empty()                      # fiksēts lapas augšā, redzams ritinot

            def _show_banner(text, color):
                banner.markdown(
                    "<div style='position:fixed;top:0;left:0;right:0;z-index:1000000;padding:0.6rem 1rem;"
                    f"text-align:center;font-weight:600;color:#fff;background:{color};"
                    "box-shadow:0 2px 8px rgba(0,0,0,.3)'>" + text + "</div>",
                    unsafe_allow_html=True)

            _show_banner(T["live_provisional"], "#d97706")      # analīze vēl notiek visu laiku

            def _show_partial(md, _stage):
                live.markdown(linkify(md))

            compare_data = None
            try:
                with st.spinner(T["spinner_analysis"]):
                    if compare_mode:
                        # 3 neatkarīgas analīzes (kešs apiets, nonce katrai atšķirīgs), rezultāts netiek keširāts
                        base_nonce = datetime.datetime.now().timestamp()
                        runs = []
                        for _i in range(COMPARE_RUNS):
                            banner.empty()
                            _show_banner(f"{T['live_provisional']} ({_i + 1}/{COMPARE_RUNS})", "#d97706")
                            runs.append(analyze_contract(
                                api_key, T["report_lang_name"], kind, content, base_nonce + _i + 1,
                                progress=_show_partial, store=False,
                            ))
                        analysis_all, laws_source = runs[0]
                        compare_data = ae.compare_reports([r[0] for r in runs])
                        compare_data["reports"] = [r[0] for r in runs]
                        compare_data["sources"] = [r[1] for r in runs]
                    else:
                        analysis_all, laws_source = analyze_contract(
                            api_key, T["report_lang_name"], kind, content,
                            datetime.datetime.now().timestamp() if force_new else 0,
                            progress=_show_partial,
                        )
            finally:
                banner.empty()
            st.session_state["just_done"] = True
            _secs = (datetime.datetime.now() - _t0).total_seconds()
            if "💾" not in laws_source and _secs < 2:
                laws_source += " | 💾 0 cost (cache)"
            else:
                laws_source += f" | ⏱ {int(_secs // 60)} min {int(_secs % 60)} s" if _secs >= 60 else f" | ⏱ {_secs:.0f} s"
            st.caption(f"{T['law_source']}: {laws_source}")
            live.markdown(analysis_all)

            if analysis_all:
                st.markdown(sources_md())
                report_text = analysis_all + "\n\n## " + T["sources_head"] + "\n" + sources_md()
                result = build_downloads(report_text)
                result.update({
                    "analysis": analysis_all,
                    "laws_source": laws_source,
                })
                if compare_data:
                    result["compare"] = compare_data
                # Saglabā, lai analīze nepazustu pēc lejupielādes
                session_add(
                    result, st.session_state.get("terms", {}),
                    "Ielīmēts teksts" if file_extension == "txt" else uploaded_file.name, "real",
                    contract_text=content if kind == "txt" else "",
                )
                render_downloads(result)

        except Exception as e:
            st.error(f"{T['err_process']}: {str(e)}")
        if st.session_state.get("result") is not prev_result:
            st.rerun()
if "result" in st.session_state:
    # Parāda saglabāto analīzi (arī pēc lejupielādes, citu elementu izmantošanas vai kļūdas)
    saved = st.session_state["result"]
    loc = localize_analysis(saved["analysis"])              # virsraksti un pantu nosaukumi pašreizējā valodā
    report_key = T["report_lang_name"] + "|h2"              # h2: virsraksti ar paskaidrojumiem iekavās
    if saved.get("lang") != report_key:                     # valoda vai virsraksti mainīti: Word un PDF no jauna
        saved.update(build_downloads(loc + "\n\n## " + T["sources_head"] + "\n" + sources_md()))
        saved["lang"] = report_key
    with results_slot:                                      # cilnes lapas augšā
        if st.session_state.pop("just_done", False):
            st.success(T["live_done"])
        render_history_picker()
        st.markdown(
            "<style>"
            "[data-testid='stTabs'] [role='tablist']{gap:0!important;overflow:hidden!important;width:100%!important}"
            "[data-testid='stTab'],[data-testid='stTabs'] [role='tab']{"
            "flex:1 1 0!important;min-width:0!important;width:25%!important;height:auto!important;"
            "padding:0.35rem 0.1rem 0.45rem!important;justify-content:center!important;"
            "white-space:normal!important}"
            "[data-testid='stTab'] *,[data-testid='stTabs'] [role='tab'] *{white-space:normal!important;"
            "text-align:center!important;font-size:clamp(0.62rem,2.2vw,1.15rem)!important;line-height:1.15!important;"
            "overflow-wrap:anywhere;word-break:break-word}"
            "</style>", unsafe_allow_html=True)
        tab_risks, tab_chat, tab_docs, tab_calc = st.tabs([
            T.get("tab_risks", "📊 Risk analysis"),
            T.get("tab_chat", "💬 Chat with the contract"),
            T.get("tab_docs", "📄 Document download"),
            T.get("tab_calc", "🧮 Penalty calculator"),
        ])
        with tab_risks:                                     # kopsavilkums + detaļas
            st.caption(f"{T['law_source']}: {localize_source(saved['laws_source'])}")
            note = saved.get("cache_note", "")
            if note:
                st.info(T["test_note"] if note.startswith("🧪") else note)
            if saved.get("compare"):
                render_compare(saved["compare"])
            render_dashboard(loc)
            render_sections(loc, skip=("🟢", "⚪"))        # 📊 🔴 🟡 🔵
            render_sections(loc, only=("🟢",))             # 🟢 ieteiktie labojumi visiem K, B, A
            render_sections(loc, only=("⚪",))             # ⚪ neitrālās piezīmes - beigās
            render_law_help(loc)
            inject_law_tips(loc)
            st.markdown(sources_md())
        with tab_calc:
            render_calculator(boxed=False)
        with tab_docs:                                      # Word/PDF: analīze + labotais līgums
            render_downloads(saved)
            render_fixed_contract(saved["analysis"])
        with tab_chat:
            render_chat_tab()
        st.divider()
else:
    render_calculator()
