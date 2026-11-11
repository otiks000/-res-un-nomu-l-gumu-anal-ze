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
    page_title="Īres un nomas līgumu analizētājs",
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
        "mock_penalty": "🧪 **Līgumsods un procenti:** līgumsods par kavējumu kopumā nevar pārsniegt 10% no parāda {cl1716}, un tam jābūt samērīgam {cl1717}. Konkrētu summu aprēķiniet kalkulatorā lapas apakšā.",
        "mock_rent": "🧪 **Īres maksa:** līguma laikā to drīkst mainīt tikai tad, ja līgumā ir paredzēti maksas grozīšanas nosacījumi un kārtība {dtil10}.",
        "mock_repairs": "🧪 **Remonti:** līguma 6.1.2. punktā „sīkie remontdarbi” nav definēti, tāpēc var rasties strīds, kurš maksā par sadzīves tehnikas vai santehnikas bojājumiem. Izīrētājs savukārt nodrošina telpas lietošanas iespēju visu līguma laiku {dtil18}.",
        "mock_default": "🧪 **Testa atbilde:** jautājums saņemts („{question}”). Testa režīmā Gemini netiek izsaukts. Īstajā režīmā atbildi sagatavotu Gemini, balstoties uz līguma tekstu. Pamēģiniet jautāt par mājdzīvniekiem, uzteikumu, drošības naudu, zemesgrāmatu, līgumsodu, īres maksu vai remontiem.",
    },
    "English": {
        "reset_btn": "🔄 Clear session / Start over",
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
        "mock_penalty": "🧪 **Penalty and interest:** a penalty for late payment cannot in total exceed 10% of the debt {cl1716} and must be proportionate {cl1717}. Calculate the exact amount with the calculator at the bottom of the page.",
        "mock_rent": "🧪 **Rent:** during the contract it may be changed only if the contract sets out conditions and a procedure for changing it {dtil10}.",
        "mock_repairs": "🧪 **Repairs:** clause 6.1.2 of the contract does not define “minor repairs”, so a dispute may arise over who pays for damage to appliances or plumbing. The landlord, in turn, must keep the premises usable throughout the contract {dtil18}.",
        "mock_default": "🧪 **Test answer:** your question was received (“{question}”). In test mode Gemini is not called. In real mode Gemini would answer based on the contract text. Try asking about pets, termination, the security deposit, the land register, penalties, rent or repairs.",
    },
    "Русский": {
        "reset_btn": "🔄 Очистить сессию / Начать заново",
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
        "mock_penalty": "🧪 **Неустойка и проценты:** неустойка за просрочку в совокупности не может превышать 10% от долга {cl1716} и должна быть соразмерной {cl1717}. Точную сумму рассчитайте в калькуляторе внизу страницы.",
        "mock_rent": "🧪 **Арендная плата:** в течение срока договора её можно изменять только в том случае, если в договоре предусмотрены условия и порядок изменения платы {dtil10}.",
        "mock_repairs": "🧪 **Ремонт:** в пункте 6.1.2 договора «мелкий ремонт» не определён, поэтому может возникнуть спор о том, кто платит за повреждения бытовой техники или сантехники. Арендодатель, в свою очередь, обеспечивает возможность пользоваться помещением в течение всего срока договора {dtil18}.",
        "mock_default": "🧪 **Тестовый ответ:** вопрос получен («{question}»). В тестовом режиме Gemini не вызывается. В реальном режиме ответ подготовил бы Gemini на основе текста договора. Попробуйте спросить о домашних животных, расторжении, залоге, земельной книге, неустойке, арендной плате или ремонте.",
    },
}
for _l, _extra in TR_NEW.items():
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

# Sānu josla
st.sidebar.header(T["sidebar_header"])
api_key = st.sidebar.text_input(T["api_label"], type="password", key="api_key_input")
st.sidebar.button(T["reset_btn"], on_click=reset_session, use_container_width=True)
uploaded_file = st.file_uploader(
    T["uploader"],
    type=["docx", "pdf", "png", "jpg", "jpeg"],
    key=f"uploader_{st.session_state.get('uploader_gen', 0)}",
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
USE_CONTEXT_CACHE = False

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
test_mode = st.checkbox(T["test_mode_label"], key="test_mode")
force_new = st.checkbox(T["force_new_label"], key="force_new")
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
    model = genai.GenerativeModel(
        "gemini-3.8-flash",
        generation_config=genai.GenerationConfig(temperature=0.2, max_output_tokens=800),
    )
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


def render_calculator():
    """Soda naudas kalkulators izvēlētajā valodā. Ievades dati tiek saglabāti sesijā katrai analīzei atsevišķi."""
    entry = current_entry()
    terms = (entry["terms"] if entry else st.session_state.get("terms")) or {}
    sig = entry["id"] if entry else "none"
    store = st.session_state.setdefault("calc_store", {})
    saved = store.get(sig, {})
    with st.expander(T["calc_title"], expanded=False):
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


def chat_answer(question, entry, previous):
    """Reālais režīms: jautājums + līguma konteksts uz Gemini. Atbilde izvēlētajā saskarnes valodā."""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        "gemini-3.8-flash",
        generation_config=genai.GenerationConfig(temperature=0.2, max_output_tokens=800),
    )
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
    return linkify(model.generate_content("\n\n".join(parts)).text)


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
                if use_mock:
                    answer = linkify(mock_chat_answer(prompt))
                elif not api_key:
                    answer = T["chat_no_key"]
                else:
                    try:
                        with st.spinner(T["chat_spinner"]):
                            answer = chat_answer(prompt, entry, mine)
                    except Exception as e:
                        answer = tfmt("chat_error", error=e)
                st.markdown(answer)
            msgs.append({"role": "assistant", "content": answer, "entry": entry["id"], "mock": use_mock})
            mine = [m for m in msgs if m["entry"] == entry["id"]]

    components.html(CHAT_JS.replace("__TOKEN__", f"{len(mine)}-{int(is_open)}"), height=0)
    inject_law_tips(entry["result"]["analysis"] + "\n" + "\n".join(shown(m) for m in mine))
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
HEAD_EMOJI_KEYS = {"📊": "h_overall", "🔴": "h_crit", "🟡": "h_warn", "🟢": "h_recs", "🔵": "h_border", "⚪": "h_neutral"}


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


# ---------------------------------------------------------------
# Analīzes izpilde
# ---------------------------------------------------------------
# ---------------------------------------------------------------
# Testa režīms: iebūvēts parauga rezultāts, Gemini netiek izsaukts (0 €)
# ---------------------------------------------------------------
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

## 🟢 {T['h_recs']}
K1. 2.2. punkts: Pievienot punktu par īres tiesību nostiprināšanu zemesgrāmatā. {T['label_proposed']}: „Puses vienojas nostiprināt īres tiesības zemesgrāmatā. Izīrētājs 10 darba dienu laikā paraksta nostiprinājuma lūgumu.”

K2. 4.3. punkts: Atļaut ieturēt tikai ar abpusēji parakstītu defektu aktu un izdevumu apliecinošiem dokumentiem.

K3. 7.1. punkts: Noteikt kopējo līgumsoda griestu 10% no parāda.

B1. 3.1. punkts: Ierakstīt īres maksas summu cipariem un vārdiem.

B5. 7.1. punkts: Atstāt vienu mehānismu: vai nu procentus, vai līgumsodu.

## 🔵 {T['h_border']}
A1. 4.5. punkts (Depozīta atmaksa): Termiņš 1 mēnesis ir pieļaujams, bet garāks par likuma noklusējumu. [[DTIL 12]] {T['label_proposed']}: „Izīrētājs atgriež drošības naudu 10 darba dienu laikā pēc nodošanas-pieņemšanas akta parakstīšanas.”

A2. 1. sadaļa (Bankas konts): Īrniekam nav vietas bankas rekvizītiem.

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
    session_add(result, {"rent": 450.0, "rate_day": 0.5}, "Testa paraugs", "test")
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
    saved.setdefault("lang", T["report_lang_name"])
    if saved["lang"] != T["report_lang_name"]:              # valoda mainīta: Word un PDF tiek izveidoti no jauna
        saved.update(build_downloads(loc + "\n\n## " + T["sources_head"] + "\n" + sources_md()))
        saved["lang"] = T["report_lang_name"]
    render_history_picker()
    st.caption(f"{T['law_source']}: {localize_source(saved['laws_source'])}")
    note = saved.get("cache_note", "")
    if note:
        st.info(T["test_note"] if note.startswith("🧪") else note)
    render_dashboard(loc)
    render_sections(loc)
    render_law_help(loc)
    inject_law_tips(loc)
    st.markdown(sources_md())
    render_downloads(saved)
render_chat_widget()
render_calculator()
