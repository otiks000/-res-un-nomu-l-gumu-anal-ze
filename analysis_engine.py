"""Līgumu analīzes dzinējs: fiksēts kritēriju saraksts + 2 slāņu verifikācija + SQLite kešs.

Dzinējs NEATKARĪGS no Streamlit un Gemini SDK: modeļa izsaukumus tam padod kā funkcijas
(sk. AnalysisContext), tāpēc to var testēt bez API atslēgas.

Plūsma jaunam līgumam:
  1. slānis (1 izsaukums): līgums pret VISIEM CHECKLIST kritērijiem -> stingrs JSON.
     Python pārbauda, ka katrs kritērijs ir atgriezts, ka citāti tiešām ir līgumā, un
     vajadzības gadījumā (tikai trūkstošajiem) pieprasa labojumu.
  2. slānis (1 izsaukums): TIKAI atrastie riski (kritiski + brīdinājumi) pret precīziem likuma
     pantiem. Python pieņem tikai tādus pantus, kas tiešām eksistē likuma tekstā.
  Python sakārto un uzraksta gala atskaiti, tāpēc nekas netiek izlaists un secība vienmēr ir tā pati.
"""
import hashlib
import json
import re
import sqlite3
import time
import unicodedata
from dataclasses import dataclass, field
from typing import Callable

LEVELS = ("nav", "apsverams", "bridinajums", "kritisks")
RANK = {lvl: i for i, lvl in enumerate(LEVELS)}
SECTION = {"kritisks": ("🔴", "K"), "bridinajums": ("🟡", "B"), "apsverams": ("🔵", "A")}

# ---------------------------------------------------------------
# Fiksētais pārbaudes saraksts. floor = minimālais risks, ja punkts līgumā TRŪKST.
# Sarakstu un secību kontrolē Python: modelis nevar nevienu kritēriju "aizmirst".
# ---------------------------------------------------------------
CHECKLIST = [
    ("PARTIES", "Pušu dati un rekvizīti",
     "Vai ir paredzēta vieta pušu vārdiem, personas kodiem, adresēm, tālrunim, e-pastam (tukši lauki NAV kļūda; kļūda ir, ja vietas VISPĀR nav).", "bridinajums"),
    ("PREMISES", "Telpu apraksts",
     "Precīza adrese, platība, kadastra apzīmējums, telpu raksturojums (DTIL 8. pants).", "bridinajums"),
    ("TERM", "Līguma termiņš",
     "Vai ir noteikts sākuma un beigu datums (līgums jāslēdz uz noteiktu termiņu). Tukšs termiņa lauks ir jāatzīmē.", "bridinajums"),
    ("RENEWAL", "Pagarināšana",
     "Kā līgumu var pagarināt vai atjaunot; vai nav automātisku vienpusēju nosacījumu.", "apsverams"),
    ("TENANT_NOTICE", "Īrnieka uzteikums",
     "Īrnieka tiesības izbeigt līgumu, uzteikuma termiņš un kārtība (nedrīkst būt sliktāk par likumu, DTIL 20., 27. pants).", "bridinajums"),
    ("LANDLORD_NOTICE", "Izīrētāja uzteikums un atkāpšanās",
     "Pamati un termiņi, kad izīrētājs drīkst izbeigt līgumu (tikai likumā noteiktie gadījumi, DTIL 21.–27. pants); vienpusēja izlikšana.", "bridinajums"),
    ("RENT_AMOUNT", "Īres maksas apmērs",
     "Summa, valūta, kas ir iekļauts maksā. Tukša summa ir jāatzīmē.", "bridinajums"),
    ("RENT_PAYMENT", "Īres maksas samaksa",
     "Maksāšanas termiņš, kārtība, kur un kā maksāt (konts/IBAN, maksājuma saņēmējs).", "bridinajums"),
    ("RENT_CHANGE", "Īres maksas maiņa",
     "Vai izīrētājs var vienpusēji palielināt maksu; vai ir likumiski nosacījumi un kārtība (DTIL 10. pants).", "bridinajums"),
    ("UTILITIES", "Komunālie maksājumi",
     "Kas, cik un kā maksā komunālos pakalpojumus; rēķinu uzrādīšanas kārtība, skaitītāji, norēķini.", "bridinajums"),
    ("FUND", "Uzkrājumu fonds un kapitālie maksājumi",
     "Vai īrniekam uzlikti kapitālā remonta, uzkrājumu fonda vai citi īpašnieka pienākumi apmaksāt maksājumi.", "nav"),
    ("DEPOSIT_AMOUNT", "Drošības depozīta apmērs",
     "Apmērs; nedrīkst pārsniegt divu mēnešu īres maksu (DTIL 12. pants).", "apsverams"),
    ("DEPOSIT_USE", "Depozīta ieturēšana",
     "Kādos gadījumos un kā drīkst ieturēt; vai tiek paziņots un pierādīti zaudējumi; patvaļīga ieturēšana.", "apsverams"),
    ("DEPOSIT_REFUND", "Depozīta atmaksa",
     "Atmaksas termiņš (garāks par 10 darba dienām = apsverams punkts), kārtība.", "bridinajums"),
    ("LATE_INTEREST", "Nokavējuma procenti",
     "Procentu apmērs un aprēķins; vai nav pārmērīgi (CL 1764.–1765. pants).", "apsverams"),
    ("PENALTY_CAP", "Līgumsodi un to griesti",
     "Līgumsoda apmērs, griesti, samērīgums (CL 1716., 1717. pants); neskaidri 'VAI' formulējumi starp soda veidiem.", "bridinajums"),
    ("LIABILITY", "Atbildība par zaudējumiem",
     "Atbildība bez vainas vai bez pierādījumiem, pārmērīga atbildība par bojājumiem, atbildības ierobežojumi.", "nav"),
    ("REPAIRS_MINOR", "Sīkie remontdarbi",
     "Vai ir definēts, kas ir 'sīkie remontdarbi' un kurš tos sedz.", "bridinajums"),
    ("REPAIRS_MAJOR", "Lielie remonti un avārijas",
     "Kas sedz lielos remontus, avārijas un apkopi; izīrētāja pienākumi uzturēt telpas (DTIL 18. pants).", "bridinajums"),
    ("ENTRY", "Izīrētāja piekļuve un apskate",
     "Vai izīrētājs drīkst ienākt mājoklī; brīdinājuma laiks, biežums, iemesli (DTIL 13. pants). Brīvas ieiešanas tiesības = risks.", "bridinajums"),
    ("HANDOVER", "Pieņemšanas-nodošanas akts",
     "Akts, telpu un inventāra stāvoklis nodošanas brīdī un atdošanas brīdī.", "bridinajums"),
    ("OCCUPANTS", "Iemitinātās personas un deklarēšana",
     "Personu skaits, ģimenes locekļi, tiesības deklarēt dzīvesvietu (DTIL 14., 15. pants).", "apsverams"),
    ("RESTRICTIONS", "Apakšīre, dzīvnieki un lietošanas ierobežojumi",
     "Apakšīre, mājdzīvnieki, viesi, smēķēšana, citi aizliegumi un to sekas.", "apsverams"),
    ("LAND_REGISTRY", "Zemesgrāmata un īpašnieka maiņa",
     "Vai paredzēta īres tiesību nostiprināšana zemesgrāmatā; bez tās jaunais īpašnieks var prasīt atbrīvot telpas (DTIL 28., 35. pants).", "kritisks"),
    ("VACATE", "Telpu atbrīvošana",
     "Telpu atbrīvošanas kārtība un termiņi līguma beigās; soda sankcijas par aizkavēšanos (DTIL 29. pants).", "apsverams"),
    ("DISPUTES", "Strīdu risināšana un piemērojamais likums",
     "Piemērojamais likums, jurisdikcija, strīdu kārtība, paziņojumu forma un adreses.", "bridinajums"),
    ("BLANKS", "Tukši vai neskaidri būtiski lauki",
     "Finansiāli vai juridiski būtiski tukši lauki (summas, termiņi, procenti) un nenoteikti termini ('citi maksājumi', 'pēc vajadzības').", "nav"),
]
CHECK_IDS = [c[0] for c in CHECKLIST]
CHECK_BY_ID = {c[0]: c for c in CHECKLIST}

LAYER1_SYSTEM = """Tu esi pieredzējis Latvijas īres tiesību eksperts. Pārbaudi dotos līgumu pret FIKSĒTU kritēriju sarakstu un atgriezi TIKAI derīgu JSON, bez cita teksta.

KLASIFIKĀCIJA (riska_limenis):
- "kritisks": punkts pārkāpj likuma imperatīvās normas (Dzīvojamo telpu īres likums, Civillikums) vai rada nepamatotus finansiālus zaudējumus (patvaļīga depozīta ieturēšana, nesamērīgs līgumsods, vienpusēja īres maksas palielināšana bez likumā noteiktā pamata, atbildība bez vainas, kapitālā remonta maksājumi uz īrnieka rēķina, īres tiesību nenostiprināšana zemesgrāmatā).
- "bridinajums": punkts nav skaidrs vai trūkst (neskaidrs formulējums, trūkstošs termiņš, summa vai kārtība, 'VAI' starp sodiem), bet tas nav tiešs pārkāpums.
- "apsverams": likumīgs, bet vienai pusei neizdevīgs vai nestandarta punkts (piemēram, depozīta atmaksa ilgāk par 10 darba dienām). Ja šaubies, vai tā ir problēma, NEIZLAID, bet liec šeit.
- "nav": kritērijs izpildīts un problēmu nav.

TUKŠIE LAUKI: tukši lauki, kas jāaizpilda parakstot un kas ir personas vai identifikācijas dati (vārdi, personas kodi, adreses, tālruņi, e-pasti, bankas konti, kadastra numuri, paraksti, noslēgšanas vieta un datums), NAV kļūda, un tos neatzīmē. Finansiāli vai juridiski būtiski tukši lauki (īres summa, depozīta summa, termiņi, procenti, uzteikuma termiņš) JĀATZĪMĒ. Ja līgumā VISPĀR nav vietas svarīgiem datiem (piem., kur maksāt īri), tas jāatzīmē.

NOTEIKUMI:
- Atgriez PRECĪZI vienu ierakstu katram kritērija id, tādā secībā kā sarakstā. Neizlaid nevienu id.
- "atrasts": true, ja līgumā ir punkts par šo tēmu; false, ja tēmas līgumā nav. Ja tēmas nav, bet likums to prasa, riska_limenis nedrīkst būt zemāks par kritērijā norādīto "ja trūkst" līmeni.
- "originalais_teksts": precīzs, burtisks citāts no līguma (kopē rakstzīmi pa rakstzīmei, līdz 300 rakstzīmēm; ar "..." vari izlaist vidu). Ja atrasts=false, atstāj "".
- "punkts": tikai līguma punkta numurs (piemēram "4.3"), ja zināms, citādi "".
- Ierakstiem ar riska_limenis "nav" visas teksta vērtības atstāj "" (tas ietaupa izmaksas).
- Katra atsevišķa problēma ir ATSEVIŠĶS ieraksts; nekad nesavieno vairākus punktus. Ja kritērijā ir vairākas atsevišķas problēmas, ieraksti tās masīvā "papildus".
- "papildus" ir paredzēts ARĪ riskiem, kas neietilpst nevienā kritērijā (id ievieto "CITS"). Ja tādu nav, "papildus": [].
- "ko_mainit": īss darbības apraksts. Ja punkts līgumā trūkst, sāc ar "Pievienot" (vai tā tulkojumu atskaites valodā).
- "labojums": gatavs piedāvātais līguma teksts līguma valodā, BEZ pēdiņu zīmēm („ " “ ” « »). Tukšs, ja nav nepieciešams.
- "labojums_tulkojums": labojuma tulkojums atskaites valodā; tukšs, ja līguma valoda sakrīt ar atskaites valodu.
- Esi rūpīgs un kritisks: labāk atzīmē vairāk potenciālu problēmu nekā mazāk.
- "kopsavilkums": 1–2 teikumi atskaites valodā: vai līgumu drīkst parakstīt un kāds ir galvenais risks.

JSON STRUKTŪRA:
{"kopsavilkums": "...", "items": [{"id": "...", "atrasts": true, "riska_limenis": "nav|apsverams|bridinajums|kritisks", "punkts": "", "nosaukums": "", "originalais_teksts": "", "problema": "", "ko_mainit": "", "labojums": "", "labojums_tulkojums": ""}], "papildus": [ same structure with "id": "CITS" ]}
"""

# Statiskā daļa (valodas neatkarīga): to var droši ievietot Gemini Context Cache kopā ar likumu tekstu.
LAYER2_SYSTEM = """Tu esi Latvijas īres tiesību eksperts un tavs uzdevums ir PĀRBAUDĪT jau atrastus līguma riska punktus pret precīziem likumu pantiem.
Tev ir dots oficiālais likumu teksts un saraksts ar atrastiem riska punktiem. Atgriez TIKAI derīgu JSON.

Katram punktam (pēc id):
- "pamatojums": 1 teikums, kurš pants (vai kurš finansiālais zaudējums) apstiprina vai noraida risku.
- "verdikts":
    "apstiprinats" - punkts tiešām pārkāpj konkrētu likuma pantu VAI rada nepamatotus finansiālus zaudējumus;
    "pazeminat_bridinajums" - nav skaidra pārkāpuma, bet punkts ir neskaidrs vai nepilnīgs;
    "pazeminat_apsverams" - punkts ir likumīgs, bet vienai pusei neizdevīgs.
  Punktiem, kas atzīmēti kā brīdinājums, vienmēr raksti "apstiprinats" (tev tikai jāatrod panti).
- "panti": masīvs ar pantiem, kas šo punktu tieši regulē: [{"likums": "DTIL" vai "CL", "nr": panta numurs}]. Drīksti norādīt TIKAI pantus, kas ir sniegtajā likumu tekstā. Ja neesi pārliecināts, atstāj tukšu. Neizdomā pantu numurus.

JSON STRUKTŪRA:
{"kopsavilkums": "...", "items": [{"id": "...", "pamatojums": "...", "verdikts": "...", "panti": []}]}
"kopsavilkums": 1–2 teikumi atskaites valodā, kuros ņemti vērā TIKAI pēc pārbaudes palikušie kritiskie riski: vai līgumu drīkst parakstīt un kāds ir galvenais risks.
Atbildes tekstus ("pamatojums", "kopsavilkums") raksti atskaites valodā.
"""
LAYER2_LAW_HEADER = (
    "SPĒKĀ ESOŠIE LIKUMI (OFICIĀLAIS TEKSTS):\n"
    "Balsties tikai uz šiem tekstiem. Atsaucies uz pantiem tikai tad, ja tie ir šajā tekstā.\n\n"
)

# Mainot kritērijus vai promptus, atslēga mainās automātiski -> vecie keša ieraksti vairs netiek izmantoti.
PROMPT_FINGERPRINT = hashlib.sha256(
    json.dumps([CHECKLIST, LAYER1_SYSTEM, LAYER2_SYSTEM], ensure_ascii=False).encode("utf-8")
).hexdigest()[:10]


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
    labels: dict                   # T vārdnīca: h_overall, h_crit, h_warn, h_border, h_recs, label_proposed, label_translation
    llm: Callable                  # llm(parts:list, max_tokens:int) -> (teksts, (ievade_žetoni, izvade_žetoni))
    llm_verify: Callable = None    # llm_verify(system_text:str, parts:list, max_tokens:int) -> tāpat; pievieno likumu tekstu (var izmantot kešu)
    valid_articles: Callable = None  # valid_articles("DTIL"|"CL") -> kopa ar esošajiem pantu numuriem
    legal_notes: str = ""          # īss pamatnoteikumu saraksts 1. slānim
    tokens_in: int = 0
    tokens_out: int = 0
    notes: list = field(default_factory=list)
    complete: bool = True          # False, ja kāds slānis neizdevās -> rezultātu nekešo


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


def _level(v):
    s = _strip_diacritics(str(v or "")).strip().lower()
    for lvl, keys in (("kritisks", ("kritisk", "critical")), ("bridinajums", ("bridinaj", "warning")),
                      ("apsverams", ("apsverams", "borderline", "consider")), ("nav", ("nav", "none", "ok"))):
        if any(s.startswith(k) for k in keys):
            return lvl
    return None


def clause_key(punkts):
    m = re.search(r"\d+(?:\.\d+)*", str(punkts or ""))
    return tuple(int(x) for x in m.group().split(".")) if m else (10 ** 6,)


def _clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _no_quotes(s):
    return _clean(re.sub(r"[„“”«»\"]", "", str(s or "")))


def _normalize_item(raw, forced_id=None):
    if not isinstance(raw, dict):
        return None
    cid = forced_id or _clean(raw.get("id")).upper()
    lvl = _level(raw.get("riska_limenis"))
    if lvl is None:
        return None
    return {
        "id": cid,
        "atrasts": _bool(raw.get("atrasts")),
        "level": lvl,
        "punkts": _clean(raw.get("punkts")),
        "nosaukums": _clean(raw.get("nosaukums")),
        "quote": _clean(raw.get("originalais_teksts")),
        "problema": _clean(raw.get("problema")),
        "ko_mainit": _clean(raw.get("ko_mainit")),
        "labojums": _no_quotes(raw.get("labojums")),
        "tulkojums": _clean(raw.get("labojums_tulkojums")),
        "articles": [],
    }


def _ask(ctx, parts, system_fn=None, max_tokens=16000, tries=2):
    """Izsauc modeli un parsē JSON; vienreiz atkārto, ja JSON nav derīgs."""
    last = None
    for _ in range(tries):
        text, usage = (system_fn or ctx.llm)(parts, max_tokens)
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


def _checklist_text(ids):
    lines = []
    for cid in ids:
        _, name, what, floor = CHECK_BY_ID[cid]
        lines.append(f'- id="{cid}" | {name} | pārbaudi: {what} | ja trūkst, vismaz: {floor}')
    return "\n".join(lines)


def _layer1_intro(ctx, ids):
    return (
        LAYER1_SYSTEM
        + f"\nATSKAITES VALODA: {ctx.lang_name}. Visus tekstus (nosaukums, problema, ko_mainit, kopsavilkums, labojums_tulkojums) raksti šajā valodā. "
          "Citātus ("
          "originalais_teksts) un labojumu raksti līguma oriģinālvalodā.\n"
        + (f"\nPAMATNOTEIKUMI (kopsavilkums no likuma):\n{ctx.legal_notes}\n" if ctx.legal_notes else "")
        + "\nKRITĒRIJU SARAKSTS (atgriez ierakstu katram id):\n" + _checklist_text(ids)
    )


# ---------------------------------------------------------------
# 1. slānis
# ---------------------------------------------------------------
def _problems(item, contract_norm):
    """Atgriež iemeslu, kāpēc ieraksts jālabo, vai None, ja viss kārtībā."""
    if item["level"] == "nav":
        return None
    if not item["problema"]:
        return "tukša problēma"
    cid = item["id"]
    floor = CHECK_BY_ID[cid][3] if cid in CHECK_BY_ID else "nav"
    if not item["atrasts"] and RANK[item["level"]] < RANK[floor]:
        return f"trūkstošam punktam risks nedrīkst būt zemāks par {floor}"
    if item["atrasts"] and contract_norm is not None and not quote_in_contract(item["quote"], contract_norm):
        return "citāts nav atrodams līgumā (jābūt burtiskam)"
    return None


def run_layer1(ctx):
    contract_norm = _norm_text(ctx.content) if ctx.kind == "txt" else None
    data = _ask(ctx, _content_parts(ctx, _layer1_intro(ctx, CHECK_IDS)))
    summary = _clean(data.get("kopsavilkums"))
    by_id, extras = {}, []

    def absorb(payload):
        for raw in payload.get("items", []) or []:
            it = _normalize_item(raw)
            if it and it["id"] in CHECK_BY_ID and it["id"] not in by_id:
                by_id[it["id"]] = it
        for raw in payload.get("papildus", []) or []:
            it = _normalize_item(raw, forced_id="CITS")
            if it and it["level"] != "nav" and it["problema"]:
                extras.append(it)

    absorb(data)

    def todo():
        bad = {}
        for cid in CHECK_IDS:
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

    unchecked = []
    for cid, why in todo().items():
        it = by_id.get(cid)
        if it is None:
            unchecked.append(cid)
        elif "citāts" in why:
            it["quote"] = ""                  # pazaudēt neesošu citātu, bet pašu atradumu paturēt
        elif "zemāks" in why:
            it["level"] = CHECK_BY_ID[cid][3]  # Python nosaka minimumu trūkstošiem obligātiem punktiem
        elif "problēma" in why:
            it["level"] = "nav"
    items = [by_id[c] for c in CHECK_IDS if c in by_id]
    for e in extras:                           # papildu atradumus pārbauda tāpat (citāts)
        if contract_norm is not None and e["quote"] and not quote_in_contract(e["quote"], contract_norm):
            e["quote"] = ""
    return items, extras, summary, unchecked


# ---------------------------------------------------------------
# 2. slānis: tikai atrastie riski pret likuma pantiem
# ---------------------------------------------------------------
def run_layer2(ctx, findings):
    if not findings or ctx.llm_verify is None:
        return None
    payload = [
        {"id": f["_uid"], "riska_limenis": "kritisks" if f["level"] == "kritisks" else "bridinajums",
         "punkts": f["punkts"], "nosaukums": f["nosaukums"], "originalais_teksts": f["quote"], "problema": f["problema"]}
        for f in findings if f["level"] in ("kritisks", "bridinajums")
    ]
    if not payload:
        return None
    dynamic = (f"ATSKAITES VALODA: {ctx.lang_name}.\n\nPĀRBAUDĀMIE PUNKTI (JSON):\n"
               + json.dumps(payload, ensure_ascii=False))

    def call(parts, max_tokens):
        return ctx.llm_verify(LAYER2_SYSTEM, parts, max_tokens)

    return _ask(ctx, [dynamic], system_fn=call)


def apply_layer2(ctx, findings, data):
    """Pielieto verifikāciju. Nekad neizdzēš atradumu: kritisko drīkst tikai pazemināt."""
    verified = 0
    by_uid = {f["_uid"]: f for f in findings}
    valid = {code: (ctx.valid_articles(code) if ctx.valid_articles else set()) for code in ("DTIL", "CL")}
    for raw in data.get("items", []) or []:
        f = by_uid.get(_clean(raw.get("id")))
        if not f:
            continue
        verified += 1
        arts = []
        for a in raw.get("panti", []) or []:
            try:
                code, nr = str(a.get("likums", "")).upper(), int(a.get("nr"))
            except (AttributeError, TypeError, ValueError):
                continue
            if code in valid and nr in valid[code] and (code, nr) not in arts:
                arts.append((code, nr))
        f["articles"] = arts[:4]
        verdict = _strip_diacritics(str(raw.get("verdikts", ""))).lower()
        if f["level"] == "kritisks":
            if "apsverams" in verdict:
                f["level"] = "apsverams"
            elif "bridinaj" in verdict:
                f["level"] = "bridinajums"
    return verified


# ---------------------------------------------------------------
# Atskaites veidošana (Python): fiksēta secība, kodi K/B/A, ieteikumi visiem kodiem
# ---------------------------------------------------------------
def render_report(findings, summary, labels):
    order = {cid: i for i, cid in enumerate(CHECK_IDS)}
    groups = {lvl: [] for lvl in SECTION}
    for f in findings:
        if f["level"] in groups:
            groups[f["level"]].append(f)
    for lst in groups.values():
        lst.sort(key=lambda f: (clause_key(f["punkts"]), order.get(f["id"], 999), f["problema"]))

    heads = {"kritisks": labels["h_crit"], "bridinajums": labels["h_warn"], "apsverams": labels["h_border"]}
    out = [f"## 📊 {labels['h_overall']}", summary or "", ""]
    coded = []
    for lvl in ("kritisks", "bridinajums", "apsverams"):
        emoji, letter = SECTION[lvl]
        out.append(f"## {emoji} {heads[lvl]}")
        if not groups[lvl]:
            out += ["–", ""]
            continue
        for i, f in enumerate(groups[lvl], 1):
            code = f"{letter}{i}"
            coded.append((code, f))
            ref = _ref(f)
            cites = "".join(f" [[{c} {n}]]" for c, n in f["articles"])
            out.append(f"{code}. {ref}: {f['problema'].rstrip()}{cites}")
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
    return "\n".join(out).strip() + "\n"


def _ref(f, with_title=True):
    m = re.search(r"\d+(?:\.\d+)*", f["punkts"] or "")
    num = (m.group() + ".") if m else ""
    title = f["nosaukums"] or (CHECK_BY_ID[f["id"]][1] if f["id"] in CHECK_BY_ID else "")
    if with_title:
        return f"{num} ({title})".strip() if num else title
    return num or title


# ---------------------------------------------------------------
# Galvenā funkcija
# ---------------------------------------------------------------
def run_analysis(ctx):
    """Atgriež (markdown atskaite, info rinda). ctx.complete=False, ja kāds slānis neizdevās (nekešot)."""
    items, extras, summary, unchecked = run_layer1(ctx)
    findings = [f for f in items + extras if f["level"] != "nav"]
    for i, f in enumerate(findings):
        f["_uid"] = f"F{i}"
    sent = sum(1 for f in findings if f["level"] in ("kritisks", "bridinajums"))
    verified = 0
    if sent:
        try:
            data = run_layer2(ctx, findings)
            if data:
                verified = apply_layer2(ctx, findings, data)
                new_summary = _clean(data.get("kopsavilkums"))
                if new_summary:
                    summary = new_summary
            else:
                ctx.notes.append("nav likumu teksta")
        except Exception as e:                  # 1. slāņa rezultāts tiek parādīts, bet netiek kešots
            ctx.complete = False
            ctx.notes.append(f"2. slānis neizdevās: {str(e)[:60]}")
    report = render_report(findings, summary, ctx.labels)
    info = f"✅ {len(CHECK_IDS) - len(unchecked)}/{len(CHECK_IDS)} | 🔎 {verified}/{sent} | 🪙 {ctx.tokens_in}↑ {ctx.tokens_out}↓"
    if unchecked:
        ctx.complete = False
        info += " | ⚠️ " + ",".join(unchecked)
    if ctx.notes:
        info += " | " + "; ".join(ctx.notes)
    return report, info
