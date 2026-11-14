"""Līgumu analīzes dzinējs: fiksēts kritēriju saraksts, objektīva klasifikācija, SQLite kešs.

Dzinējs NEATKARĪGS no Streamlit un Gemini SDK: modeļa izsaukumus tam padod kā funkcijas
(sk. AnalysisContext), tāpēc to var testēt bez API atslēgas.

Plūsma jaunam līgumam:
  1. slānis: kritēriji sadalīti grupās (nauda / termiņi un izbeigšana / pārējie); katra grupa
     redz VISU līgumu un paralēli atgriež stingru JSON. Modelis nevērtē, tas tikai nosaka faktus:
     veidu, likuma pantus un jā/nē pazīmes. Riska līmeni (🔴 🟡 💶 ⚪) piešķir Python pēc noteikumiem.
  2. Papildu pārbaude (sweep): atrod to, ko grupas izlaida.
  3. Python pārbauda skaitliskās robežas, pantu esamību, citātus un uzraksta atskaiti.
"""
import hashlib
import json
import re
import sqlite3
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import Callable

SECTION = {"kritisks": ("🔴", "K"), "bridinajums": ("🟡", "B"), "nauda": ("💶", "M"), "neitrals": ("⚪", "N")}
KINDS = ("likuma_parkapums", "naudas_nosacijums", "neskaidrs_vai_truksts", "citi_nosacijums", "nav")
FLAGS = ("tukss", "bez_griestiem", "vienpusejs", "bez_vainas")   # objektīvas naudas pazīmes (jā/nē)
_LOCK = threading.Lock()

# ---------------------------------------------------------------
# Fiksētais pārbaudes saraksts. floor = minimālais risks, ja punkts līgumā TRŪKST.
# Sarakstu un secību kontrolē Python: modelis nevar nevienu kritēriju "aizmirst".
# ---------------------------------------------------------------
CHECKLIST = [
    ("PARTIES", "Pušu dati un rekvizīti",
     "Vai ir paredzēta vieta pušu vārdiem, personas kodiem, adresēm, tālrunim, e-pastam (tukši lauki NAV kļūda; kļūda ir, ja vietas VISPĀR nav).", "apsverams"),
    ("PREMISES", "Telpu apraksts",
     "Precīza adrese, platība, kadastra apzīmējums, telpu raksturojums (DTIL 8. pants).", "kritisks"),
    ("TERM", "Līguma termiņš",
     "Vai ir noteikts sākuma un beigu datums (līgums jāslēdz uz noteiktu termiņu). Tukšs termiņa lauks ir jāatzīmē.", "kritisks"),
    ("RENEWAL", "Pagarināšana",
     "Kā līgumu var pagarināt vai atjaunot; vai nav automātisku vienpusēju nosacījumu.", "apsverams"),
    ("TENANT_NOTICE", "Īrnieka uzteikums",
     "Īrnieka tiesības izbeigt līgumu, uzteikuma termiņš un kārtība (nedrīkst būt sliktāk par likumu, DTIL 20., 27. pants).", "kritisks"),
    ("LANDLORD_NOTICE", "Izīrētāja uzteikums un atkāpšanās",
     "Pamati un termiņi, kad izīrētājs drīkst izbeigt līgumu (tikai likumā noteiktie gadījumi, DTIL 21.–27. pants); vienpusēja izlikšana.", "kritisks"),
    ("RENT_AMOUNT", "Īres maksas apmērs",
     "Summa, valūta, kas ir iekļauts maksā. Tukša summa ir jāatzīmē.", "kritisks"),
    ("RENT_PAYMENT", "Īres maksas samaksa",
     "Maksāšanas termiņš, kārtība, kur un kā maksāt (konts/IBAN, maksājuma saņēmējs).", "apsverams"),
    ("RENT_CHANGE", "Īres maksas maiņa",
     "Vai izīrētājs var vienpusēji palielināt maksu; vai ir likumiski nosacījumi un kārtība (DTIL 10. pants).", "kritisks"),
    ("UTILITIES", "Komunālie maksājumi",
     "Kas, cik un kā maksā komunālos pakalpojumus; rēķinu uzrādīšanas kārtība, skaitītāji, norēķini.", "kritisks"),
    ("FUND", "Uzkrājumu fonds un kapitālie maksājumi",
     "Vai īrniekam uzlikti kapitālā remonta, uzkrājumu fonda vai citi īpašnieka pienākumi apmaksāt maksājumi.", "nav"),
    ("DEPOSIT_AMOUNT", "Drošības depozīta apmērs",
     "Apmērs; nedrīkst pārsniegt divu mēnešu īres maksu (DTIL 12. pants).", "apsverams"),
    ("DEPOSIT_USE", "Depozīta ieturēšana",
     "Kādos gadījumos un kā drīkst ieturēt; vai tiek paziņots un pierādīti zaudējumi; patvaļīga ieturēšana.", "kritisks"),
    ("DEPOSIT_REFUND", "Depozīta atmaksa",
     "Atmaksas termiņš (garāks par 10 darba dienām = apsverams punkts), kārtība.", "kritisks"),
    ("LATE_INTEREST", "Nokavējuma procenti",
     "Procentu apmērs un aprēķins; vai nav pārmērīgi (CL 1764.–1765. pants).", "nav"),
    ("PENALTY_CAP", "Līgumsodi un to griesti",
     "Līgumsoda apmērs, griesti, samērīgums (CL 1716., 1717. pants); neskaidri 'VAI' formulējumi starp soda veidiem.", "nav"),
    ("LIABILITY", "Atbildība par zaudējumiem",
     "Atbildība bez vainas vai bez pierādījumiem, pārmērīga atbildība par bojājumiem, atbildības ierobežojumi.", "nav"),
    ("REPAIRS_MINOR", "Sīkie remontdarbi",
     "Vai ir definēts, kas ir 'sīkie remontdarbi' un kurš tos sedz.", "kritisks"),
    ("REPAIRS_MAJOR", "Lielie remonti un avārijas",
     "Kas sedz lielos remontus, avārijas un apkopi; izīrētāja pienākumi uzturēt telpas (DTIL 18. pants).", "kritisks"),
    ("ENTRY", "Izīrētāja piekļuve un apskate",
     "Vai izīrētājs drīkst ienākt mājoklī; brīdinājuma laiks, biežums, iemesli (DTIL 13. pants). Brīvas ieiešanas tiesības = risks.", "apsverams"),
    ("HANDOVER", "Pieņemšanas-nodošanas akts",
     "Akts, telpu un inventāra stāvoklis nodošanas brīdī un atdošanas brīdī.", "apsverams"),
    ("OCCUPANTS", "Iemitinātās personas un deklarēšana",
     "Personu skaits, ģimenes locekļi, tiesības deklarēt dzīvesvietu (DTIL 14., 15. pants).", "apsverams"),
    ("RESTRICTIONS", "Apakšīre, dzīvnieki un lietošanas ierobežojumi",
     "Apakšīre, mājdzīvnieki, viesi, smēķēšana, citi aizliegumi un to sekas.", "apsverams"),
    ("LAND_REGISTRY", "Zemesgrāmata un īpašnieka maiņa",
     "Vai paredzēta īres tiesību nostiprināšana zemesgrāmatā; bez tās jaunais īpašnieks var prasīt atbrīvot telpas (DTIL 28., 35. pants).", "apsverams"),
    ("VACATE", "Telpu atbrīvošana",
     "Telpu atbrīvošanas kārtība un termiņi līguma beigās; soda sankcijas par aizkavēšanos (DTIL 29. pants).", "apsverams"),
    ("DISPUTES", "Strīdu risināšana un piemērojamais likums",
     "Piemērojamais likums, jurisdikcija, strīdu kārtība, paziņojumu forma un adreses.", "apsverams"),
    ("VAT", "PVN",
     "Vai ir norādīts, vai īres/nomas maksa ir ar PVN vai bez tā un kura puse maksā PVN (tukšs vai nenoteikts = tukss).", "kritisks"),
    ("USE_PURPOSE", "Atļautais lietošanas veids",
     "Kādam mērķim un kādai darbībai telpas drīkst izmantot; ierobežojumi (CL 2151. pants).", "apsverams"),
    ("IMPROVEMENTS", "Uzlabojumi un pārbūve",
     "Vai nomnieks drīkst veikt uzlabojumus vai pārbūvi (ar kādu piekrišanu), kam tie pieder un vai tos atlīdzina līguma beigās (CL 2140., 2155., 2156. pants).", "apsverams"),
    ("INSURANCE", "Apdrošināšana",
     "Kura puse apdrošina telpas, īpašumu un civiltiesisko atbildību.", "apsverams"),
    ("PERMITS", "Atļaujas un drošība",
     "Vai paredzētas darbībai nepieciešamās atļaujas, ugunsdrošības un citas prasības un kura puse par tām atbild.", "apsverams"),
    ("SIGNAGE", "Reklāma un izkārtnes",
     "Vai nomnieks drīkst izvietot izkārtnes un reklāmu; kārtība.", "nav"),
    ("BLANKS", "Tukši vai neskaidri būtiski lauki",
     "Finansiāli vai juridiski būtiski tukši lauki (summas, termiņi, procenti) un nenoteikti termini ('citi maksājumi', 'pēc vajadzības').", "nav"),
]
MONEY_IDS = {"VAT", "RENT_AMOUNT", "RENT_PAYMENT", "RENT_CHANGE", "UTILITIES", "FUND", "DEPOSIT_AMOUNT", "DEPOSIT_USE",
             "DEPOSIT_REFUND", "LATE_INTEREST", "PENALTY_CAP", "LIABILITY", "REPAIRS_MINOR", "REPAIRS_MAJOR"}
# Panti, kas skar naudu un zaudējumus: to pārkāpums ir 🔴 (tikai naudas kritērijiem); visi pārējie likuma pārkāpumi ir 🟡.
# DTIL 8 (līguma saturs) šeit nav: modelis to min gandrīz pie jebkura trūkstoša lauka, un tas pārslēgtu līmeni starp skrējieniem.
MONEY_ARTICLES = {("DTIL", 10), ("DTIL", 11), ("DTIL", 12), ("DTIL", 18),
                  ("CL", 1716), ("CL", 1717), ("CL", 1764), ("CL", 1765)}
CHECK_IDS = [c[0] for c in CHECKLIST]
CHECK_BY_ID = {c[0]: c for c in CHECKLIST}
# Trūkstoša kritērija līmenis (fiksēta tabula; modelis to neietekmē). None = trūkums netiek atzīmēts.
MISSING_LEVEL = {c[0]: {"kritisks": "kritisks", "apsverams": "bridinajums", "nav": None}[c[3]] for c in CHECKLIST}
# Kuram līguma veidam kritērijs attiecas: dz = dzīvojamā īre, noma = nedzīvojamo telpu noma/cits; pārējie abiem.
APPLIES = {"OCCUPANTS": "dz", "VAT": "noma", "USE_PURPOSE": "noma", "IMPROVEMENTS": "noma",
           "INSURANCE": "noma", "PERMITS": "noma", "SIGNAGE": "noma"}
HARD_BLANK_IDS = MONEY_IDS | {"TERM"}      # tukša vai nenoteikta vērtība šeit ir 🔴
# Ne-naudas kritēriji, kuros jebkura problēma (likuma pārkāpums vai neskaidrība) ir 🔴. Fiksēts saraksts pēc līguma veida,
# lai krāsa nebūtu atkarīga no tā, vai modelis punktu nosauca par "pārkāpumu" vai par "neskaidrību".
LAW_RED_IDS = {
    "dz": {"TERM", "TENANT_NOTICE", "LANDLORD_NOTICE", "ENTRY", "LAND_REGISTRY"},   # DTIL: termiņš, uzteikums, netraucēta lietošana, zemesgrāmata
    "noma": {"TENANT_NOTICE", "LANDLORD_NOTICE", "LAND_REGISTRY"},                  # CL: uzteikums, patvarīga izlikšana, zemesgrāmata (2174. p.)
}
TYPE_LABEL = {"dz": "dzīvojamo telpu īre", "noma": "telpu noma", "cits": "cits līgums"}
TYPE_FROM_MODEL = {"dzivojama_ire": "dz", "telpu_noma": "noma", "cits": "cits"}


def _group(tip):
    return "dz" if tip == "dz" else "noma"


def applies(cid, tip):
    """Vai kritērijs attiecas uz līguma veidu tip ('dz'/'noma'/'cits'; tukšs = nav zināms, der visi)."""
    return not tip or APPLIES.get(cid, "abi") in ("abi", _group(tip))


def detect_contract_type(text):
    """Līguma veids pēc atslēgvārdiem bez modeļa izsaukuma: 'dz', 'noma' vai '' (nav skaidrs)."""
    t = _strip_diacritics(str(text or "").lower())
    non_res = t.count("nedzivojam")
    res = t.count("dzivojam") - non_res + t.count("dzivokl")
    n_dz = res + t.count("irnieks") + t.count("izireta")
    n_noma = 4 * non_res + t.count("nomnieks") + t.count("iznomata")
    if non_res and non_res >= res:
        return "noma"
    if not n_dz and not n_noma:
        return ""
    if n_noma >= 2 * n_dz:
        return "noma"
    if n_dz >= 2 * n_noma:
        return "dz"
    return ""
# Kritēriju grupas: katra grupa redz visu līgumu, bet pārbauda šaurāku tēmu loku (stabilāks rezultāts).
CHECK_GROUPS = [
    [c for c in CHECK_IDS if c in MONEY_IDS or c == "BLANKS"],
    ["PARTIES", "PREMISES", "TERM", "RENEWAL", "TENANT_NOTICE", "LANDLORD_NOTICE", "LAND_REGISTRY", "VACATE", "DISPUTES"],
    ["ENTRY", "HANDOVER", "OCCUPANTS", "RESTRICTIONS", "USE_PURPOSE", "IMPROVEMENTS", "INSURANCE", "PERMITS", "SIGNAGE"],
]
assert sorted(sum(CHECK_GROUPS, [])) == sorted(CHECK_IDS)

LAYER1_SYSTEM = """Tu esi pieredzējis Latvijas īres un nomas tiesību eksperts. Pārbaudi doto līgumu pret DOTAJIEM kritērijiem un atgriez TIKAI derīgu JSON, bez cita teksta.

MĒRĶIS: atrast VISUS punktus, kas attiecas uz dotajiem kritērijiem, neviena neizlaižot. Izej cauri līgumam punkts pa punktam (katru numurēto punktu un apakšpunktu) un katram piemēro visus dotos kritērijus.

TU NEIZSAKI VIEDOKLI UN NEVĒRTĒ, VAI KAS IR "NEIZDEVĪGS". Tu nosaki tikai faktus: veidu, likuma pantus un jā/nē pazīmes. Riska līmeni piešķir programma.

LĪGUMA VEIDS ("lig_veids"): "dzivojama_ire" (dzīvojamo telpu īre), "telpu_noma" (nedzīvojamo telpu vai citas lietas noma) vai "cits". Nosaki pēc līguma satura un norādi JSON laukā "lig_veids". Kritērijus ar atzīmi [TIKAI DZĪVOJAMAI ĪREI] vai [TIKAI NOMAI] piemēro tikai attiecīgajam veidam; citam veidam to ierakstu atgriez ar veidu "nav" un atrasts=false. Nomas līgumam Dzīvojamo telpu īres likumu (DTIL) NEPIEMĒRO un tā pantus NEMINI.

VEIDS ("veids"), izvēlies VIENU:
- "likuma_parkapums": punkts pārkāpj konkrētu likuma pantu. Norādi to laukā "panti" (tikai panti no PAMATNOTEIKUMIEM). Ja nevari nosaukt konkrētu pantu, šo veidu NELIETO.
- "naudas_nosacijums": punkts nosaka naudas maksājumu vai naudas sekas (īre, depozīts, komunālie, soda nauda, procenti, remonti, atbildība par bojājumiem) un NAV likuma pārkāpums. "problema" = neitrāls FAKTU apraksts (kas maksā, cik, līdz kuram termiņam, uz kāda pamata), bez vērtējuma un bez vārdiem "neizdevīgs", "pārāk", "nesamērīgs".
- "neskaidrs_vai_truksts": nenoteikti termini, pretrunas starp punktiem, nepilnīgi pienākumi, trūkstošas procedūras vai tēma, kuras līgumā NAV (atrasts=false), ja tas nav naudas nosacījums.
- "citi_nosacijums": likumīgs nosacījums, kas nav par naudu, bet ir atzīmējams (piem., apakšīres aizliegums, viesu ierobežojumi, apskates tiesības). "problema" = neitrāls apraksts.
- "nav": kritērijs izpildīts, atzīmēt nav ko. Visas teksta vērtības atstāj "".
Valodas, gramatikas un pareizrakstības kļūdas šeit NEVĒRTĒ (tās izvērtē cits modulis).

NAUDAS PAZĪMES ("pazimes", jā/nē, tikai pēc līguma teksta; katram naudas atradumam): 
- "tukss": summa, termiņš vai procents ir tukšs vai nenoteikts (arī "citi maksājumi", "pēc vajadzības").
- "bez_griestiem": līgumsods, procenti vai maksājums bez augšējās robežas vai bez summas.
- "vienpusejs": izīrētājs var vienpusēji mainīt maksu vai nosacījumus, ieturēt naudu bez paziņošanas un pierādījumiem vai vienpusēji izlikt.
- "bez_vainas": īrnieks atbild par zaudējumiem bez vainas vai bez pierādījumiem.

CENAI JĀBŪT FIKSĒTAI: īres vai nomas maksa, depozīts, soda procents un termiņi jābūt ierakstīti uzreiz. Tukšs lauks (_____), "pēc vienošanās", "tirgus apstākļi" vai "pēc vajadzības" ir pazīme "tukss". Tas neattiecas uz personas datiem zemāk.
TUKŠIE LAUKI: tukši lauki, kas jāaizpilda parakstot un kas ir personas vai identifikācijas dati (vārdi, personas kodi, adreses, tālruņi, e-pasti, bankas konti, kadastra numuri, paraksti, noslēgšanas vieta un datums), NAV kļūda, un tos neatzīmē. Finansiāli vai juridiski būtiski tukši lauki (īres summa, depozīta summa, termiņi, procenti, uzteikuma termiņš) jāatzīmē (pazīme "tukss"). Ja līgumā VISPĀR nav vietas svarīgiem datiem (piem., kur maksāt īri), tas jāatzīmē.

NOTEIKUMI:
- Atgriez vismaz vienu ierakstu katram dotajam kritērija id, tādā secībā kā sarakstā. Neizlaid nevienu id.
- VAIRĀKAS PROBLĒMAS: ja kritērijam ir vairākas atsevišķas problēmas vai nosacījumi (dažādos punktos vai dažādi pārkāpumi vienā punktā), atgriez katru kā ATSEVIŠĶU ierakstu ar to pašu id. Nekad nesavieno vairākus punktus vienā ierakstā un nekad neierobežo atradumu skaitu.
- "atrasts": true, ja līgumā ir punkts par šo tēmu; false, ja tēmas līgumā nav. Trūkstošai tēmai veids ir "neskaidrs_vai_truksts" (vai naudas gadījumā "naudas_nosacijums" ar pazīmi "tukss").
- "originalais_teksts": precīzs, burtisks citāts no līguma (kopē rakstzīmi pa rakstzīmei, līdz 300 rakstzīmēm; ar "..." vari izlaist vidu). Ja atrasts=false, atstāj "".
- "punkts": tikai līguma punkta numurs (piemēram "4.3"), ja zināms, citādi "".
- "panti": masīvs [{"likums": "DTIL" vai "CL", "nr": skaitlis}] ar pantiem, ko punkts pārkāpj vai kas to regulē. Tikai panti no PAMATNOTEIKUMIEM; ja neesi pārliecināts, atstāj []. Neizdomā pantu numurus.
- "problema": konkrēti, 1–2 teikumi: likuma pārkāpumam kurš pants un kāpēc; pārējiem veidiem neitrāls fakts vai kas nav izrunāts.
- "ko_mainit": īss darbības apraksts. Ja punkts līgumā trūkst, sāc ar "Pievienot" (vai tā tulkojumu atskaites valodā). Tikai veidiem "likuma_parkapums" un "neskaidrs_vai_truksts".
- "labojums": gatavs piedāvātais līguma teksts līguma valodā, BEZ pēdiņu zīmēm („ " “ ” « »). Tukšs, ja nav nepieciešams.
- "labojums_tulkojums": labojuma tulkojums atskaites valodā; tukšs, ja līguma valoda sakrīt ar atskaites valodu.
- "papildus" ir paredzēts atradumiem, kas neietilpst nevienā dotajā kritērijā (id "CITS"). Ja tādu nav, "papildus": [].
- "kopsavilkums": 1–2 teikumi atskaites valodā: vai līgumu drīkst parakstīt un kāds ir galvenais risks.
- "skaitli": skaitļi, ko izvelc burtiski no līguma (nekad neaprēķini un neizdomā; ja nav norādīts, null): "ires_summa_mened", "depozits", "depozita_atmaksa_darba_dienas", "soda_likme_diena_proc" (līgumsoda vai nokavējuma procenti, pārrēķināti uz % dienā; null, ja nav), "tuksi_naudas_lauki" (true, ja ir tukša vai nenoteikta īres, depozīta, soda vai procentu summa/termiņš), "neierobezoti_maksajumi" (true, ja ir maksājumi bez summas vai griestiem).

JSON STRUKTŪRA:
{"lig_veids": "dzivojama_ire|telpu_noma|cits", "kopsavilkums": "...", "skaitli": {"ires_summa_mened": null, "depozits": null, "depozita_atmaksa_darba_dienas": null, "soda_likme_diena_proc": null, "tuksi_naudas_lauki": false, "neierobezoti_maksajumi": false}, "items": [{"id": "...", "atrasts": true, "veids": "likuma_parkapums|naudas_nosacijums|neskaidrs_vai_truksts|citi_nosacijums|nav", "pazimes": {"tukss": false, "bez_griestiem": false, "vienpusejs": false, "bez_vainas": false}, "panti": [], "punkts": "", "nosaukums": "", "originalais_teksts": "", "problema": "", "ko_mainit": "", "labojums": "", "labojums_tulkojums": ""}], "papildus": [ same structure with "id": "CITS" ]}
"""

# ---------------------------------------------------------------
# SQLite kešs
# ---------------------------------------------------------------
def content_hash(content):
    """SHA-256 no līguma teksta (normalizēta atstarpes) vai faila baitiem."""
    if isinstance(content, str):
        content = re.sub(r"\s+", " ", content).strip().encode("utf-8")
    return hashlib.sha256(content).hexdigest()


class AnalysisCache:
    """Noturīgs kešs: (līguma SHA-256, prompta versija, modelis, atskaites valoda) -> gatava analīze."""

    def __init__(self, path="analysis_cache.db"):
        self.path = path
        with self._conn() as c:
            c.execute(
                "CREATE TABLE IF NOT EXISTS analyses ("
                "content_hash TEXT NOT NULL, prompt_ver TEXT NOT NULL, model TEXT NOT NULL, lang TEXT NOT NULL,"
                "payload TEXT NOT NULL, created_at REAL NOT NULL, hits INTEGER NOT NULL DEFAULT 0,"
                "PRIMARY KEY (content_hash, prompt_ver, model, lang))"
            )

    def _conn(self):
        conn = sqlite3.connect(self.path, timeout=15)
        try:
            conn.execute("PRAGMA journal_mode=WAL")
        except sqlite3.DatabaseError:
            pass
        return conn

    def get(self, chash, prompt_ver, model, lang):
        try:
            with self._conn() as c:
                row = c.execute(
                    "SELECT payload FROM analyses WHERE content_hash=? AND prompt_ver=? AND model=? AND lang=?",
                    (chash, prompt_ver, model, lang),
                ).fetchone()
                if not row:
                    return None
                c.execute(
                    "UPDATE analyses SET hits=hits+1 WHERE content_hash=? AND prompt_ver=? AND model=? AND lang=?",
                    (chash, prompt_ver, model, lang),
                )
            return json.loads(row[0])
        except (sqlite3.Error, ValueError):
            return None            # bojāts kešs nedrīkst apturēt analīzi

    def put(self, chash, prompt_ver, model, lang, payload):
        try:
            with self._conn() as c:
                c.execute(
                    "INSERT OR REPLACE INTO analyses (content_hash, prompt_ver, model, lang, payload, created_at, hits) "
                    "VALUES (?,?,?,?,?,?,0)",
                    (chash, prompt_ver, model, lang, json.dumps(payload, ensure_ascii=False), time.time()),
                )
        except sqlite3.Error:
            pass

    def clear(self):
        with self._conn() as c:
            c.execute("DELETE FROM analyses")


# ---------------------------------------------------------------
# Konteksts un palīgfunkcijas
# ---------------------------------------------------------------
@dataclass
class AnalysisContext:
    kind: str                      # 'txt' vai faila paplašinājums (pdf/png/jpg)
    content: object                # teksts vai baiti
    lang_name: str                 # atskaites valoda ("latviešu", "angļu"...)
    labels: dict                   # T vārdnīca: h_overall, h_crit, h_warn, h_money, h_neutral, h_recs, label_proposed, label_translation
    llm: Callable                  # llm(parts:list, max_tokens:int) -> (teksts, (ievade_žetoni, izvade_žetoni))
    valid_articles: Callable = None  # valid_articles("DTIL"|"CL") -> kopa ar esošajiem pantu numuriem
    legal_notes: str = ""          # īss pamatnoteikumu saraksts 1. slānim
    legal_notes_noma: str = ""     # tas pats nomas līgumiem (bez DTIL); ja tukšs, lieto legal_notes
    contract_type: str = ""        # 'dz' | 'noma' | 'cits' | '' (nav zināms; to nosaka detect_contract_type vai modelis)
    check_ids: list = field(default_factory=list)   # kritēriji, kas attiecas uz šo līgumu
    tokens_in: int = 0
    tokens_out: int = 0
    notes: list = field(default_factory=list)
    complete: bool = True          # False, ja kāds slānis neizdevās -> rezultātu nekešo
    figures: dict = field(default_factory=dict)   # skaitļi, ko modelis izvilcis no līguma (robežas pārbauda Python)
    neutral_count: int = 0         # cik neitrālo piezīmju (⚪) ir atskaitē (korektūra turpina numerāciju)
    progress: Callable = None      # progress(markdown_vai_None, statusa_teksts): starprezultāts lietotāja saskarnei (neietekmē analīzi)


def _strip_diacritics(s):
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")


def _norm_text(s):
    return re.sub(r"[\W_]+", " ", str(s).lower()).strip()


def quote_in_contract(quote, contract_norm):
    """Vai citāts (ar iespējamiem '...' izlaidumiem) burtiski ir līgumā (ignorējot atstarpes un pieturzīmes)."""
    parts = [p for p in re.split(r"\.\.\.|…", str(quote)) if len(_norm_text(p)) >= 8]
    return bool(parts) and all(_norm_text(p) in contract_norm for p in parts)


def _parse_json(text):
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t).strip()
    try:
        return json.loads(t)
    except ValueError:
        m = re.search(r"\{.*\}", t, re.S)
        if not m:
            raise
        return json.loads(m.group())


def _bool(v):
    if isinstance(v, str):
        return _strip_diacritics(v).strip().lower() in ("true", "jā", "ja", "yes", "1", "atrasts")
    return bool(v)


def clause_key(punkts):
    m = re.search(r"\d+(?:\.\d+)*", str(punkts or ""))
    return tuple(int(x) for x in m.group().split(".")) if m else (10 ** 6,)


def _clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _no_quotes(s):
    return _clean(re.sub(r"[„“”«»\"]", "", str(s or "")))


def _kind(v):
    s = _strip_diacritics(str(v or "")).strip().lower()
    return s if s in KINDS else None


def _normalize_item(raw, forced_id=None):
    if not isinstance(raw, dict):
        return None
    cid = forced_id or _clean(raw.get("id")).upper()
    kind = _kind(raw.get("veids"))
    if kind is None:
        return None
    flags_raw = raw.get("pazimes") if isinstance(raw.get("pazimes"), dict) else {}
    return {
        "id": cid,
        "atrasts": _bool(raw.get("atrasts")),
        "kind": kind,
        "flags": {k: _bool(flags_raw.get(k)) for k in FLAGS},
        "articles_raw": raw.get("panti") if isinstance(raw.get("panti"), list) else [],
        "punkts": _clean(raw.get("punkts")),
        "nosaukums": _clean(raw.get("nosaukums")),
        "quote": _clean(raw.get("originalais_teksts")),
        "problema": _clean(raw.get("problema")),
        "ko_mainit": _clean(raw.get("ko_mainit")),
        "labojums": _no_quotes(raw.get("labojums")),
        "tulkojums": _clean(raw.get("labojums_tulkojums")),
        "articles": [],
        "force_red": False,
        "level": "nav",
    }


def _validate_articles(ctx, f):
    """Atstāj tikai pantus, kas tiešām eksistē likuma tekstā."""
    valid = {code: (ctx.valid_articles(code) if ctx.valid_articles else None) for code in ("DTIL", "CL")}
    arts = []
    for a in f.get("articles_raw", []) or []:
        try:
            code, nr = str(a.get("likums", "")).upper(), int(a.get("nr"))
        except (AttributeError, TypeError, ValueError):
            continue
        if code == "DTIL" and ctx.contract_type in ("noma", "cits"):
            continue
        if code in valid and (valid[code] is None or nr in valid[code]) and (code, nr) not in arts:
            arts.append((code, nr))
    f["articles"] = arts[:4]


def classify(f, tip="dz"):
    """Riska līmeni nosaka Python pēc fiksētas tabulas (kritērijs + jā/nē pazīmes), nevis modeļa izvēlēti panti.
    🔴: tukša/nenoteikta cena, termiņš vai procents, vienpusēja vai neierobežota nauda, naudas kritērija problēma,
    trūkstoša obligāta tēma. 🟡: citu kritēriju problēmas un trūkumi. 💶: naudas fakti. ⚪: neitrāli nosacījumi."""
    if f["kind"] == "nav":
        return "nav"
    cid = f["id"]
    money_item = cid in MONEY_IDS or f["kind"] == "naudas_nosacijums"
    if f.get("force_red"):
        return "kritisks"
    if (money_item and any(f["flags"].values())) or (cid in HARD_BLANK_IDS and f["flags"]["tukss"]):
        return "kritisks"                       # objektīvi naudas gadījumi
    if not f["atrasts"]:                        # trūkstošs punkts: līmeni nosaka kritērija tabula
        return (MISSING_LEVEL.get(cid) or "nav") if cid in CHECK_BY_ID else "bridinajums"
    if f["kind"] == "naudas_nosacijums":
        return "nauda"
    if f["kind"] == "citi_nosacijums":
        return "neitrals"
    if cid in CHECK_BY_ID:                      # likuma pārkāpums vai neskaidrība: naudas un likuma kritēriji 🔴, pārējie 🟡
        return "kritisks" if cid in MONEY_IDS or cid in LAW_RED_IDS[_group(tip)] else "bridinajums"
    if f["kind"] == "likuma_parkapums" and any(a in MONEY_ARTICLES for a in f["articles"]):
        return "kritisks"                       # papildu atradumi bez kritērija: naudas pants
    return "bridinajums"


def _ask(ctx, parts, max_tokens=16000, tries=2):
    """Izsauc modeli un parsē JSON; vienreiz atkārto, ja JSON nav derīgs."""
    last = None
    for _ in range(tries):
        text, usage = ctx.llm(parts, max_tokens)
        with _LOCK:
            ctx.tokens_in += usage[0]
            ctx.tokens_out += usage[1]
        try:
            return _parse_json(text)
        except ValueError as e:
            last = e
    raise ValueError(f"Modelis neatgrieza derīgu JSON: {last}")


def _content_parts(ctx, intro):
    if ctx.kind == "txt":
        return [intro + "\n\nLĪGUMA TEKSTS:\n" + ctx.content]
    mime = {"pdf": "application/pdf", "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(ctx.kind, "application/octet-stream")
    return [intro, {"mime_type": mime, "data": ctx.content}]


def _checklist_text(ids, tip=""):
    lines = []
    for cid in ids:
        _, name, what, floor = CHECK_BY_ID[cid]
        if _group(tip) == "noma" and tip:          # nomai DTIL norādes tiek izņemtas
            what = re.sub(r"\s*\([^()]*DTIL[^()]*\)", "", what)
        tag = ""
        if not tip and APPLIES.get(cid, "abi") != "abi":
            tag = " | [TIKAI DZĪVOJAMAI ĪREI]" if APPLIES[cid] == "dz" else " | [TIKAI NOMAI]"
        lines.append(f'- id="{cid}" | {name} | pārbaudi: {what}{tag}' + (" | ja trūkst, atzīmē obligāti" if MISSING_LEVEL[cid] else ""))
    return "\n".join(lines)


def _notes_for(ctx):
    return (ctx.legal_notes_noma or ctx.legal_notes) if ctx.contract_type in ("noma", "cits") else ctx.legal_notes


def _layer1_intro(ctx, ids):
    return (
        LAYER1_SYSTEM
        + f"\nATSKAITES VALODA: {ctx.lang_name}. Visus tekstus (nosaukums, problema, ko_mainit, kopsavilkums, labojums_tulkojums) raksti šajā valodā. "
          "Citātus (originalais_teksts) un labojumu raksti līguma oriģinālvalodā.\n"
        + (f"\nLĪGUMA VEIDS (jau noteikts): {TYPE_LABEL[ctx.contract_type]}.\n" if ctx.contract_type else "")
        + (f"\nPAMATNOTEIKUMI (kopsavilkums no likuma):\n{_notes_for(ctx)}\n" if _notes_for(ctx) else "")
        + "\nDOTIE KRITĒRIJI (atgriez ierakstu katram id; citus kritērijus šajā izsaukumā nepārbaudi):\n" + _checklist_text(ids, ctx.contract_type)
    )


# ---------------------------------------------------------------
# 1. slānis: kritēriju grupas (paralēli)
# ---------------------------------------------------------------
def _problems(item, contract_norm):
    """Atgriež iemeslu, kāpēc ieraksts jālabo, vai None, ja viss kārtībā."""
    if item["kind"] == "nav":
        return None
    if not item["problema"]:
        return "tukša problēma"
    if item["atrasts"] and contract_norm is not None and not quote_in_contract(item["quote"], contract_norm):
        return "citāts nav atrodams līgumā (jābūt burtiskam)"
    return None


MISSING_TEXT = {
    "latviešu": "Šis punkts līgumā nav atrasts.",
    "angļu": "This clause was not found in the contract.",
    "krievu": "Этот пункт в договоре не найден.",
}


def _run_group(ctx, ids, contract_norm):
    data = _ask(ctx, _content_parts(ctx, _layer1_intro(ctx, ids)))
    by_id, extras = {}, []

    def absorb(payload):
        for raw in payload.get("items", []) or []:
            it = _normalize_item(raw)
            if not it or it["id"] not in ids:
                continue
            if it["id"] not in by_id:
                by_id[it["id"]] = it
            elif it["kind"] != "nav" and it["problema"]:
                extras.append(it)          # vēl viena atsevišķa problēma tam pašam kritērijam
        for raw in payload.get("papildus", []) or []:
            it = _normalize_item(raw, forced_id="CITS")
            if it and it["kind"] != "nav" and it["problema"]:
                extras.append(it)

    absorb(data)

    def todo():
        bad = {}
        for cid in ids:
            if cid not in by_id:
                bad[cid] = "kritērijs nav atgriezts"
            else:
                why = _problems(by_id[cid], contract_norm)
                if why:
                    bad[cid] = why
        return bad

    bad = todo()
    if bad:                                   # viens mērķēts labojums tikai problemātiskajiem kritērijiem
        intro = _layer1_intro(ctx, list(bad)) + "\n\nIEPRIEKŠĒJĀ ATBILDĒ ŠIE KRITĒRIJI BIJA NEDERĪGI. Atgriez tos vēlreiz (tikai šos id; 'papildus': []):\n" + \
            "\n".join(f"- {cid}: {why}" for cid, why in bad.items())
        try:
            fix = _ask(ctx, _content_parts(ctx, intro))
            for raw in fix.get("items", []) or []:
                it = _normalize_item(raw)
                if it and it["id"] in bad and _problems(it, contract_norm) is None:
                    by_id[it["id"]] = it
        except ValueError:
            ctx.notes.append("labošanas izsaukums neizdevās")

    for it in by_id.values():                 # obligāts punkts nav līgumā, bet modelis to nav atzīmējis: Python to atzīmē pats
        if it["kind"] == "nav" and not it["atrasts"] and MISSING_LEVEL.get(it["id"]):
            it["kind"] = "neskaidrs_vai_truksts"
            it["problema"] = MISSING_TEXT.get(ctx.lang_name, MISSING_TEXT["angļu"])
    unchecked = []
    for cid, why in todo().items():
        it = by_id.get(cid)
        if it is None:
            unchecked.append(cid)
        elif "citāts" in why:
            it["quote"] = ""                  # pazaudēt neesošu citātu, bet pašu atradumu paturēt
        elif "problēma" in why:
            it["kind"] = "nav"
    return by_id, extras, _clean(data.get("kopsavilkums")), data.get("skaitli"), unchecked, _clean(data.get("lig_veids")).lower()


def _vote_type(votes):
    """Līguma veids pēc modeļa atbildēm (vairākums; neskaidrības gadījumā dzīvojamā īre)."""
    tally = {}
    for v in votes:
        t = TYPE_FROM_MODEL.get(_strip_diacritics(v).replace(" ", "_"))
        if t:
            tally[t] = tally.get(t, 0) + 1
    return max(sorted(tally), key=lambda k: tally[k]) if tally else "dz"


def run_layer1(ctx):
    contract_norm = _norm_text(ctx.content) if ctx.kind == "txt" else None
    if not ctx.contract_type and ctx.kind == "txt":
        ctx.contract_type = detect_contract_type(ctx.content)         # bez modeļa izsaukuma
    groups = [[c for c in g if applies(c, ctx.contract_type)] for g in CHECK_GROUPS]
    groups = [g for g in groups if g]
    ctx.check_ids = [c for g in groups for c in g]
    with ThreadPoolExecutor(max_workers=len(groups)) as ex:
        futures = [ex.submit(_run_group, ctx, ids, contract_norm) for ids in groups]
        results = [f.result() for f in futures]      # secība fiksēta; kļūda grupā pārtrauc analīzi
    if not ctx.contract_type:                        # veids nebija zināms (pdf/attēls vai neskaidrs teksts): nosaka modelis
        ctx.contract_type = _vote_type([r[5] for r in results])
        ctx.check_ids = [c for c in ctx.check_ids if applies(c, ctx.contract_type)]
    by_id, extras, unchecked, summary, figures = {}, [], [], "", {}
    for b, e, summ, fig, unch, _veids in results:
        by_id.update(b)
        extras += e
        unchecked += unch
        summary = summary or summ
        for k, v in (fig.items() if isinstance(fig, dict) else []):
            if figures.get(k) in (None, "", False) and v not in (None, ""):
                figures[k] = v
    ctx.figures = figures
    items = [by_id[c] for c in CHECK_IDS if c in by_id and applies(c, ctx.contract_type)]
    unchecked = [c for c in unchecked if applies(c, ctx.contract_type)]
    extras = [e for e in extras if e["id"] == "CITS" or applies(e["id"], ctx.contract_type)]
    extras = _clean_extras(extras, contract_norm)
    return items, extras, summary, unchecked, contract_norm


def _clean_extras(extras, contract_norm):
    """Papildu atradumiem pārbauda citātu un izmet dublikātus (tas pats punkts + tas pats citāts/problēma)."""
    seen, out = set(), []
    for e in extras:
        if contract_norm is not None and e["quote"] and not quote_in_contract(e["quote"], contract_norm):
            e["quote"] = ""
        key = (e["punkts"], _norm_text(e["quote"])[:60] or _norm_text(e["problema"])[:60])
        if key in seen:
            continue
        seen.add(key)
        out.append(e)
    return sorted(out, key=lambda e: (clause_key(e["punkts"]), _norm_text(e["problema"])))


SWEEP_SYSTEM = """PAPILDU PĀRBAUDE. Tu esi Latvijas īres tiesību eksperts. Tev ir dots līgums un saraksts ar atradumiem, kas jau atrasti. Tavs uzdevums ir atrast visu, kas VĒL IR IZLAISTS. Izej cauri līgumam punkts pa punktam.
Meklē: (1) likuma pārkāpumus (ar konkrētu pantu); (2) naudas nosacījumus un naudas pazīmes (tukšs, bez griestiem, vienpusējs, bez vainas); (3) neskaidrus formulējumus, pretrunas starp punktiem un trūkstošus svarīgus punktus; (4) citus likumīgus atzīmējamus nosacījumus.
Tu NEVĒRTĒ, vai kas ir "neizdevīgs", tu nosaki tikai faktus. Neatkārto jau atrastos atradumus un neatzīmē valodas kļūdas. Tukši personas dati (vārdi, personas kodi, adreses, konti, paraksti, vieta un datums) nav kļūda.
Atgriez TIKAI derīgu JSON: {"papildus": [ieraksti]}. Katrs ieraksts: {"id": "CITS" vai tuvākais kritērija id, "atrasts": true, "veids": "likuma_parkapums|naudas_nosacijums|neskaidrs_vai_truksts|citi_nosacijums", "pazimes": {"tukss": false, "bez_griestiem": false, "vienpusejs": false, "bez_vainas": false}, "panti": [{"likums": "DTIL vai CL", "nr": 0}], "punkts": "", "nosaukums": "", "originalais_teksts": "burtisks citāts līdz 300 rakstzīmēm", "problema": "", "ko_mainit": "", "labojums": "gatavs teksts līguma valodā bez pēdiņām", "labojums_tulkojums": ""}.
Ja nekā jauna nav, atgriez {"papildus": []}."""


# Mainot kritērijus vai promptus, atslēga mainās automātiski -> vecie keša ieraksti vairs netiek izmantoti.
PROMPT_FINGERPRINT = hashlib.sha256(
    json.dumps([CHECKLIST, CHECK_GROUPS, sorted(MONEY_ARTICLES), LAYER1_SYSTEM, SWEEP_SYSTEM], ensure_ascii=False).encode("utf-8")
).hexdigest()[:10]


def run_sweep(ctx, items, extras, contract_norm):
    """Viens papildu izsaukums: atrod to, ko 1. slānis izlaida. Kļūme nav kritiska."""
    known = [{"punkts": f["punkts"], "tēma": f["nosaukums"] or f["id"], "problēma": f["problema"][:140]}
             for f in items + extras if f["kind"] != "nav"]
    intro = (SWEEP_SYSTEM + f"\nATSKAITES VALODA: {ctx.lang_name}. Tekstus raksti šajā valodā; citātus un labojumu līguma valodā.\n"
             + (f"\nLĪGUMA VEIDS: {TYPE_LABEL[ctx.contract_type]}. Nomas līgumam DTIL nepiemēro.\n" if ctx.contract_type else "")
             + (f"\nPAMATNOTEIKUMI (kopsavilkums no likuma):\n{_notes_for(ctx)}\n" if _notes_for(ctx) else "")
             + "\nJAU ATRASTIE ATRADUMI (JSON):\n" + json.dumps(known, ensure_ascii=False))
    try:
        data = _ask(ctx, _content_parts(ctx, intro))
    except ValueError:
        ctx.notes.append("papildu pārbaude neizdevās")
        return []
    new = []
    for raw in data.get("papildus", []) or []:
        it = _normalize_item(raw, forced_id="CITS")
        if it and it["kind"] != "nav" and it["problema"]:
            new.append(it)
    if not new:
        return []
    mine = {id(x) for x in new}
    return [e for e in _clean_extras(extras + new, contract_norm) if id(e) in mine]


# ---------------------------------------------------------------
# Naudas robežas: lēmumus pieņem Python, nevis modelis (stabili rezultāti starp skrējieniem)
# ---------------------------------------------------------------
DEPOSIT_MAX_MONTHS = 2          # DTIL 12. pants
REFUND_MAX_WORKDAYS = 10
RATE_MAX_PER_DAY = 0.1          # % dienā; augstāk = nesamērīgi (CL 1716., 1717., 1764. pants)

RULE_TEXT = {
    "latviešu": {
        "deposit": "Drošības depozīts ({a:g}) pārsniedz {m} mēnešu īres maksu ({b:g}).",
        "refund": "Depozīta atmaksas termiņš ({a:g} darba dienas) pārsniedz {m} darba dienas.",
        "rate": "Līgumsoda vai nokavējuma procents ({a:g}% dienā) pārsniedz {m}% dienā.",
        "blank": "Līgumā ir tukša vai nenoteikta naudas summa, termiņš vai procents.",
        "open": "Līgumā ir maksājumi bez summas vai griestiem.",
    },
    "angļu": {
        "deposit": "The security deposit ({a:g}) exceeds {m} months' rent ({b:g}).",
        "refund": "The deposit refund period ({a:g} working days) exceeds {m} working days.",
        "rate": "The penalty or late-payment rate ({a:g}% per day) exceeds {m}% per day.",
        "blank": "The contract has an empty or undefined money amount, deadline or rate.",
        "open": "The contract has payments without an amount or a cap.",
    },
    "krievu": {
        "deposit": "Залог ({a:g}) превышает арендную плату за {m} мес. ({b:g}).",
        "refund": "Срок возврата залога ({a:g} рабочих дней) превышает {m} рабочих дней.",
        "rate": "Неустойка или пеня ({a:g}% в день) превышает {m}% в день.",
        "blank": "В договоре есть пустая или неопределённая сумма, срок или процент.",
        "open": "В договоре есть платежи без суммы или предела.",
    },
}


def _num(v):
    try:
        return float(str(v).replace(",", ".").strip()) if v not in (None, "", False) else None
    except ValueError:
        return None


def apply_money_rules(ctx, items):
    """Skaitliskās robežas: pārkāpums padara atradumu par 🔴 (vai pievieno jaunu). Atgriež jaunos atradumus."""
    txt = RULE_TEXT.get(ctx.lang_name, RULE_TEXT["angļu"])
    fig = ctx.figures or {}
    by_id = {f["id"]: f for f in items}
    new = []

    def breach(cid, text):
        f = by_id.get(cid)
        if f and f["atrasts"] and f["kind"] != "nav":
            f["force_red"] = True
            if text[:40] not in f["problema"]:
                f["problema"] = (f["problema"] + " " + text).strip()
        else:
            new.append({"id": cid, "atrasts": True, "kind": "neskaidrs_vai_truksts", "flags": dict.fromkeys(FLAGS, False),
                        "articles_raw": [], "punkts": f["punkts"] if f else "", "nosaukums": CHECK_BY_ID[cid][1],
                        "quote": "", "problema": text, "ko_mainit": "", "labojums": "", "tulkojums": "",
                        "articles": [], "force_red": True, "level": "kritisks"})

    rent, dep = _num(fig.get("ires_summa_mened")), _num(fig.get("depozits"))
    resid = (ctx.contract_type or "dz") == "dz"          # depozīta un atmaksas robežas ir no DTIL: tikai dzīvojamai īrei
    if resid and rent and dep and dep > DEPOSIT_MAX_MONTHS * rent:
        breach("DEPOSIT_AMOUNT", txt["deposit"].format(a=dep, b=rent, m=DEPOSIT_MAX_MONTHS))
    days = _num(fig.get("depozita_atmaksa_darba_dienas"))
    if resid and days and days > REFUND_MAX_WORKDAYS:
        breach("DEPOSIT_REFUND", txt["refund"].format(a=days, m=REFUND_MAX_WORKDAYS))
    rate = _num(fig.get("soda_likme_diena_proc"))
    if rate and rate > RATE_MAX_PER_DAY:
        breach("PENALTY_CAP", txt["rate"].format(a=rate, m=RATE_MAX_PER_DAY))
    if _bool(fig.get("tuksi_naudas_lauki")) and not any(f["id"] in MONEY_IDS and f["level"] == "kritisks" and f["atrasts"] for f in items):
        breach("RENT_AMOUNT", txt["blank"])
    if _bool(fig.get("neierobezoti_maksajumi")) and not any(f["id"] in ("UTILITIES", "FUND") and f["level"] == "kritisks" for f in items):
        breach("UTILITIES", txt["open"])
    return new


# ---------------------------------------------------------------
# Dublikāti: viena problēma nedrīkst parādīties vairākos ierakstos (skaits svārstītos starp skrējieniem)
# ---------------------------------------------------------------
_LEVEL_RANK = {"kritisks": 3, "bridinajums": 2}


def _clause_index(content):
    """[(normalizēta rinda, punkta numurs)]: katrai līguma rindai tuvākais iepriekšējais numurs ('4.3')."""
    out, cur = [], ""
    for line in str(content or "").split("\n"):
        m = re.match(r"^\s*(\d+(?:\.\d+)*)\.?\s", line)
        if m:
            cur = m.group(1)
        n = _norm_text(line)
        if n:
            out.append((n, cur))
    return out


def assign_clauses(ctx, findings):
    """Punkta numuru no līguma teksta pēc citāta (nevis no modeļa): numurs vairs nesvārstās starp skrējieniem."""
    if ctx.kind != "txt" or not findings:
        return
    index = _clause_index(ctx.content)
    for f in findings:
        parts = [p for p in re.split(r"\.\.\.|…", f["quote"] or "") if len(_norm_text(p)) >= 8]
        if not parts:
            continue
        probe = _norm_text(parts[0])[:50]
        for n, num in index:
            if num and probe in n:
                f["punkts"] = num
                break


def merge_duplicates(findings):
    """Viena problēma = viens ieraksts: 🔴/🟡 ieraksti ar to pašu punkta numuru un to pašu līmeni tiek apvienoti
    (problēmu teksti savienoti, panti apvienoti); ieraksti bez numura tiek apvienoti pēc kritērija."""
    out, kept = [], {}
    for f in findings:
        if f["level"] not in _LEVEL_RANK:
            out.append(f)
            continue
        key = (f["level"], clause_key(f["punkts"])) if f["punkts"] else (f["level"], "id", f["id"])
        target = kept.get(key)
        if target is None:
            kept[key] = f
            out.append(f)
            continue
        extra = f["problema"].strip()
        if extra and _norm_text(extra)[:60] not in _norm_text(target["problema"]) and len(target["problema"]) < 500:
            target["problema"] = (target["problema"].rstrip() + " " + extra).strip()
        target["articles"] = (target["articles"] + [a for a in f["articles"] if a not in target["articles"]])[:4]
        if not target["labojums"] and f["labojums"]:
            for k in ("ko_mainit", "labojums", "tulkojums"):
                target[k] = f[k]
    return out


# ---------------------------------------------------------------
# Atskaites veidošana (Python): fiksēta secība, kodi K/B/M/N, ieteikumi K un B kodiem
# ---------------------------------------------------------------
def render_report(findings, summary, labels):
    order = {cid: i for i, cid in enumerate(CHECK_IDS)}
    groups = {lvl: [] for lvl in SECTION}
    for f in findings:
        if f["level"] in groups:
            groups[f["level"]].append(f)
    for lst in groups.values():
        lst.sort(key=lambda f: (clause_key(f["punkts"]), order.get(f["id"], 999), f["problema"]))

    heads = {"kritisks": labels["h_crit"], "bridinajums": labels["h_warn"],
             "nauda": labels.get("h_money", "Money terms"), "neitrals": labels.get("h_neutral", "Neutral notes")}
    out = [f"## 📊 {labels['h_overall']}", summary or "", ""]
    coded = []
    for lvl in ("kritisks", "bridinajums", "nauda"):
        emoji, letter = SECTION[lvl]
        out.append(f"## {emoji} {heads[lvl]}")
        if not groups[lvl]:
            out += ["–", ""]
            continue
        for i, f in enumerate(groups[lvl], 1):
            code = f"{letter}{i}"
            if lvl != "nauda":
                coded.append((code, f))
            cites = "".join(f" [[{c} {n}]]" for c, n in f["articles"])
            out.append(f"{code}. {_ref(f)}: {f['problema'].rstrip()}{cites}")
            if f["quote"]:
                out.append(f"> „{_no_quotes(f['quote'])}”")
            out.append("")
    out.append(f"## 🟢 {labels['h_recs']}")
    for code, f in coded:
        what = _no_quotes(f["ko_mainit"]).rstrip(" .") or _no_quotes(f["problema"]).rstrip(" .")
        line = f"{code}. {_ref(f, with_title=False)}: {what}."
        if f["labojums"]:
            line += f" {labels['label_proposed']}: „{f['labojums']}”"
        out.append(line)
        if f["tulkojums"]:
            out.append(f"{labels['label_translation']}: {f['tulkojums']}")
        out.append("")
    notes = groups["neitrals"]
    if notes:                                   # ⚪ vienmēr beigās; korektūra turpina numerāciju
        emoji, letter = SECTION["neitrals"]
        out.append(f"## {emoji} {heads['neitrals']}")
        for i, f in enumerate(notes, 1):
            out.append(f"{letter}{i}. {_ref(f)}: {f['problema'].rstrip()}")
            if f["quote"]:
                out.append(f"> „{_no_quotes(f['quote'])}”")
            out.append("")
    return "\n".join(out).strip() + "\n"


def _ref(f, with_title=True):
    m = re.search(r"\d+(?:\.\d+)*", f["punkts"] or "")
    num = (m.group() + ".") if m else ""
    title = f["nosaukums"] or (CHECK_BY_ID[f["id"]][1] if f["id"] in CHECK_BY_ID else "")
    if not num and f["id"] in CHECK_BY_ID:                # punktam bez numura nosaukums ir fiksēts (nemainās starp skrējieniem)
        title = CHECK_BY_ID[f["id"]][1]
    if with_title:
        return f"{num} ({title})".strip() if num else title
    return num or title


# ---------------------------------------------------------------
# Galvenā funkcija
# ---------------------------------------------------------------
def run_analysis(ctx):
    """Atgriež (markdown atskaite, info rinda). ctx.complete=False, ja kāds slānis neizdevās (nekešot)."""
    items, extras, summary, unchecked, contract_norm = run_layer1(ctx)
    swept = run_sweep(ctx, items, extras, contract_norm)
    allf = items + extras + swept
    dep = next((f for f in items if f["id"] == "DEPOSIT_AMOUNT"), None)
    if dep and not dep["atrasts"]:               # līgumā nav depozīta: tā ieturēšanas un atmaksas trūkums nav kļūda
        for f in items:
            if f["id"] in ("DEPOSIT_USE", "DEPOSIT_REFUND") and not f["atrasts"]:
                f["kind"] = "nav"
    assign_clauses(ctx, allf)
    for f in allf:
        _validate_articles(ctx, f)
        f["level"] = classify(f, ctx.contract_type)
    ruled = apply_money_rules(ctx, items)
    allf += ruled
    for f in allf:
        f["level"] = classify(f, ctx.contract_type)
    findings = merge_duplicates([f for f in allf if f["level"] != "nav"])
    # 💶 saraksts nedublē to, kas jau atzīmēts kā 🔴 vai 🟡 (tas pats punkta numurs vai tas pats citāts)
    flagged = [f for f in findings if f["level"] in ("kritisks", "bridinajums")]
    keys_n = {clause_key(f["punkts"]) for f in flagged if f["punkts"]}
    keys_q = {_norm_text(f["quote"])[:60] for f in flagged if f["quote"]}
    findings = [f for f in findings if not (
        f["level"] == "nauda" and ((f["punkts"] and clause_key(f["punkts"]) in keys_n)
                                   or (f["quote"] and _norm_text(f["quote"])[:60] in keys_q)))]
    ctx.neutral_count = sum(1 for f in findings if f["level"] == "neitrals")
    if ctx.progress:
        try:
            ctx.progress(render_report(findings, summary, ctx.labels), "layer1")
        except Exception:
            pass                                 # saskarnes kļūda nedrīkst traucēt analīzi
    report = render_report(findings, summary, ctx.labels)
    total = len(ctx.check_ids) or len(CHECK_IDS)
    info = (f"📄 {TYPE_LABEL.get(ctx.contract_type, '?')} | ✅ {total - len(unchecked)}/{total} (+{len(swept)}) | "
            f"🪙 {ctx.tokens_in}↑ {ctx.tokens_out}↓")
    if unchecked:
        ctx.complete = False
        info += " | ⚠️ " + ",".join(unchecked)
    if ctx.notes:
        info += " | " + "; ".join(ctx.notes)
    return report, info


# ---------------------------------------------------------------
# Neitrālā korektūra: stingrs JSON, 2 caurlaides (otrā atrod izlaisto), Python numurē un kārto
# ---------------------------------------------------------------
PROOF_SYSTEM = """Tu esi ļoti rūpīgs korektors un līgumu redaktors, kas pārzina līguma valodu (latviešu, krievu vai angļu). Atrodi VISAS valodas, rediģēšanas un neuzmanības kļūdas līgumā, pat vismazākās. Izej cauri līgumam punkts pa punktam, teikums pa teikumam. Tu nevērtē juridisko saturu un neanalizē riskus.

MEKLĒ:
- gramatikas, pareizrakstības un drukas kļūdas (garumzīmes, mīkstinājuma zīmes, locījumi, saskaņojums, dzimtes un skaitļa kļūdas);
- lieki vai trūkstoši vārdi, dubultas atstarpes, nepareizas vai trūkstošas pieturzīmes, lielo burtu lietojums;
- nekonsekventi termini vai pušu apzīmējumi (piem., "Izīrētājs" un "Iznomātājs");
- nepareizi punktu numuri, izlaisti numuri, atsauces uz neesošiem punktiem;
- pretrunas starp cipariem un vārdiem, datumiem, summām un termiņiem;
- neveikli, grūti saprotami vai divējādi lasāmi teikumi (izteikuma kļūdas).

NEATZĪMĒ: tukšus laukus, kas jāaizpilda parakstot (vārdi, personas kodi, adreses, tālruņi, e-pasti, konti, paraksti, vieta, datums); juridiskus un finansiālus riskus.

Atgriez TIKAI derīgu JSON: {"kludas": [{"punkts": "līguma punkta numurs vai ''", "kluda": "kas nepareizi (atskaites valodā)", "citats": "burtisks fragments no līguma", "labojums": "pareizais variants līguma valodā"}]}
Katra kļūda ir atsevišķs ieraksts. Ja kļūdu nav: {"kludas": []}."""


def _proof_items(data, contract_norm):
    out = []
    for raw in (data or {}).get("kludas", []) or []:
        if not isinstance(raw, dict):
            continue
        it = {"punkts": _clean(raw.get("punkts")), "kluda": _clean(raw.get("kluda")),
              "citats": _no_quotes(raw.get("citats")), "labojums": _no_quotes(raw.get("labojums"))}
        if not it["kluda"]:
            continue
        if contract_norm is not None and it["citats"] and not quote_in_contract(it["citats"], contract_norm):
            it["citats"] = ""
        out.append(it)
    return out


def run_proofread(ctx, labels, start=1, with_head=True):
    """Atgriež (markdown sadaļa ⚪, vai_izdevās). Divas caurlaides; otrā saņem jau atrastās kļūdas."""
    contract_norm = _norm_text(ctx.content) if ctx.kind == "txt" else None
    head = f"## ⚪ {labels['h_neutral']}" if with_head else ""
    base = PROOF_SYSTEM + f"\nATSKAITES VALODA (skaidrojumiem): {ctx.lang_name}."
    try:
        found = _proof_items(_ask(ctx, _content_parts(ctx, base)), contract_norm)
        known = [{"punkts": f["punkts"], "kluda": f["kluda"][:80], "citats": f["citats"][:60]} for f in found]
        more = _proof_items(_ask(ctx, _content_parts(
            ctx, base + "\n\nJAU ATRASTĀS KĻŪDAS (neatkārto tās; atgriez TIKAI to, kas vēl izlaists):\n"
            + json.dumps(known, ensure_ascii=False))), contract_norm)
    except ValueError:
        return "", False
    seen, allf = set(), []
    for f in found + more:
        key = (f["punkts"], _norm_text(f["citats"])[:40] or _norm_text(f["kluda"])[:40])
        if key not in seen:
            seen.add(key)
            allf.append(f)
    if not allf:
        return (f"{head}\n{labels.get('no_proof_errors', '–')}\n" if with_head else ""), True
    allf.sort(key=lambda f: clause_key(f["punkts"]))     # stabila kārtošana: punkta numurs, tad atrašanas secība
    out = [head] if with_head else []
    for i, f in enumerate(allf, start):
        num = (re.search(r"\d+(?:\.\d+)*", f["punkts"]) or [None])[0]
        line = f"N{i}. {num + '. ' if num else ''}{f['kluda'].rstrip()}"
        if f["citats"]:
            line += f" („{f['citats']}”)"
        if f["labojums"]:
            line += f" {labels.get('label_fix', 'Fix')}: „{f['labojums']}”."
        out.append(line)
    return "\n".join(out) + "\n", True



# ---------------------------------------------------------------
# Salīdzināšanas režīms: viena faila vairākas neatkarīgas analīzes -> kā atšķiras punkti
# ---------------------------------------------------------------
_ITEM_RE = re.compile(r"^\s*([KBM])(\d+)\.\s+(.*?):\s+(.*)$")
_LEVEL_BY_LETTER = {"K": "kritisks", "B": "bridinajums", "M": "nauda"}


def parse_report(text):
    """No gatavas markdown atskaites izvelk 🔴/🟡/💶 punktus: [{level, code, ref, key, problem, quote, fix}]."""
    items, by_code, cur, in_recs = [], {}, None, False
    for line in (text or "").split("\n"):
        if line.startswith("## "):
            in_recs = "🟢" in line
            cur = None
            continue
        m = _ITEM_RE.match(line)
        if m:
            letter, num, ref, rest = m.groups()
            code = f"{letter}{num}"
            if in_recs:                               # 🟢 ieteikums pie jau atrasta punkta
                if code in by_code:
                    fix = re.search(r"[:：]\s*„(.+?)”", rest)
                    by_code[code]["fix"] = _clean(fix.group(1) if fix else rest)
                cur = None
                continue
            if letter not in _LEVEL_BY_LETTER:
                cur = None
                continue
            title = re.search(r"\(([^)]*)\)", ref)
            ck = clause_key(ref)       # ar numuru: tikai numurs (modeļa nosaukums mainās); bez numura: fiksētais kritērija nosaukums
            key = (ck, "") if ck[0] < 10 ** 6 else (ck, _norm_text(title.group(1)) if title else _norm_text(ref))
            cur = {"level": _LEVEL_BY_LETTER[letter], "code": code, "ref": ref.strip(), "key": key,
                   "problem": _clean(re.sub(r"\[\[[^\]]*\]\]", "", rest)), "quote": "", "fix": ""}
            items.append(cur)
            by_code[code] = cur
        elif cur and line.startswith(">") and not cur["quote"]:
            cur["quote"] = _clean(line.lstrip("> "))
    return items


_SEV = {"kritisks": 3, "bridinajums": 2, "nauda": 1}


def compare_reports(reports):
    """Salīdzina N atskaites (tekstus). Atgriež:
    counts - [{level: skaits}] katrai analīzei;
    rows   - katram punktam (clause+kritērijs): ref, levels (pa analīzēm, None ja nav atrasts), problems, fixes, status;
    status: 'stable' (visās analīzēs, viens līmenis) | 'level' (visās, bet atšķirīgs līmenis) | 'partial' (ne visās);
    stability - stabilo punktu daļa 0..1."""
    n = len(reports)
    parsed = [parse_report(r) for r in reports]
    counts = [{lvl: sum(1 for it in p if it["level"] == lvl) for lvl in _LEVEL_BY_LETTER.values()} for p in parsed]
    rows = {}
    for i, p in enumerate(parsed):
        for it in p:
            row = rows.setdefault(it["key"], {"key": it["key"], "ref": it["ref"], "levels": [None] * n,
                                              "problems": [""] * n, "fixes": [""] * n, "quotes": [""] * n})
            if row["levels"][i] is None or _SEV[it["level"]] > _SEV[row["levels"][i]]:   # atkārtoti vienā analīzē - augstākais līmenis uzvar
                row["levels"][i] = it["level"]
                row["problems"][i] = it["problem"]
                row["fixes"][i] = it["fix"]
                row["quotes"][i] = it["quote"]
    out = []
    for row in sorted(rows.values(), key=lambda r: r["key"]):
        found = [l for l in row["levels"] if l]
        if len(found) < n:
            row["status"] = "partial"
        elif len(set(found)) > 1:
            row["status"] = "level"
        else:
            row["status"] = "stable"
        row["found"] = len(found)
        out.append(row)
    stable = sum(1 for r in out if r["status"] == "stable")
    return {"runs": n, "counts": counts, "rows": out, "stability": (stable / len(out)) if out else 1.0}
