import json, os, sys, tempfile, unittest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import analysis_engine as ae

CONTRACT = """1. Izīrētājs izīrē dzīvokli Rīgā, Brīvības iela 1-5.
2.2. Līgums netiek nostiprināts zemesgrāmatā.
4.3. Izīrētājs drīkst ieturēt drošības naudu pēc saviem ieskatiem.
6.1. Izīrētājs drīkst ienākt dzīvoklī jebkurā laikā."""
LABELS = {"h_overall": "Kopējais", "h_crit": "Kritiski", "h_warn": "Brīdinājumi", "h_border": "Apsveramie",
          "h_recs": "Ieteikumi", "label_proposed": "Piedāvātais formulējums", "label_translation": "Tulkojums"}


def item(cid, level="nav", **kw):
    base = {"id": cid, "atrasts": level != "nav", "riska_limenis": level, "punkts": "", "nosaukums": "",
            "originalais_teksts": "", "problema": "", "ko_mainit": "", "labojums": "", "labojums_tulkojums": ""}
    base.update(kw)
    return base


def full_items(over):
    return [over.get(cid, item(cid)) for cid in ae.CHECK_IDS]


class Fake:
    def __init__(self, l1, l1_repair=None, l2=None):
        self.l1, self.l1_repair, self.l2, self.calls = l1, l1_repair, l2, []

    def llm(self, parts, max_tokens):
        if "PAPILDU PĀRBAUDE" in str(parts):
            self.calls.append("sweep")
            return json.dumps(getattr(self, "sweep", {"papildus": []})), (100, 10)
        self.calls.append("l1")
        data = self.l1 if self.calls.count("l1") == 1 else (self.l1_repair or self.l1)
        return json.dumps(data), (100, 50)

    def verify(self, system, parts, max_tokens):
        self.calls.append("l2")
        if isinstance(self.l2, Exception):
            raise self.l2
        return json.dumps(self.l2), (500, 20)


def make_ctx(fake, kind="txt"):
    return ae.AnalysisContext(kind=kind, content=CONTRACT, lang_name="latviešu", labels=LABELS, llm=fake.llm,
                              llm_verify=fake.verify, valid_articles=lambda c: {12, 28, 35} if c == "DTIL" else {1716})


class Engine(unittest.TestCase):
    def good(self):
        return {"kopsavilkums": "Neparakstīt.", "papildus": [], "items": full_items({
            "LAND_REGISTRY": item("LAND_REGISTRY", "kritisks", punkts="2.2", nosaukums="Zemesgrāmata",
                                  originalais_teksts="Līgums netiek nostiprināts zemesgrāmatā", problema="Nav nostiprināts.",
                                  ko_mainit="Pievienot nostiprināšanu", labojums="Puses nostiprina īri zemesgrāmatā."),
            "DISPUTES": item("DISPUTES", "bridinajums", punkts="9.1", nosaukums="Strīdi",
                             originalais_teksts="jebkurā laikā", problema="Jurisdikcija nav noteikta."),
            "DEPOSIT_USE": item("DEPOSIT_USE", "kritisks", punkts="4.3", nosaukums="Depozīts",
                                originalais_teksts="drošības naudu pēc saviem ieskatiem", problema="Patvaļīga ieturēšana.", ko_mainit="Precizēt"),
            "ENTRY": item("ENTRY", "apsverams", punkts="6.1", nosaukums="Apskate",
                          originalais_teksts="jebkurā laikā", problema="Nav brīdinājuma.", ko_mainit="Pievienot 24h"),
        })}

    def test_full_run_and_layer2(self):
        l2 = {"kopsavilkums": "Pēc pārbaudes.", "items": [
            {"id": "F2", "verdikts": "pazeminat_apsverams", "panti": [{"likums": "DTIL", "nr": 28}, {"likums": "DTIL", "nr": 999}]},
            {"id": "F3", "verdikts": "pazeminat_apsverams", "panti": [{"likums": "DTIL", "nr": 12}]}]}
        f = Fake(self.good(), l2=l2)
        ctx = make_ctx(f)
        report, info = ae.run_analysis(ctx)
        self.assertEqual(f.calls, ["l1", "sweep", "l2"])
        self.assertIn("B1. 2.2. (Zemesgrāmata)", report)         # likuma pārkāpums bez naudas, pārbaude pazemināja uz 🟡
        self.assertIn("[[DTIL 28]]", report)
        self.assertNotIn("999", report)                       # neesošs pants atmests
        self.assertIn("K1. 4.3. (Depozīts)", report)           # naudas punkts paliek 🔴
        self.assertIn("B2. 6.1. (Apskate)", report)
        self.assertIn("B3. 9.1. (Strīdi)", report)
        self.assertIn("Pēc pārbaudes.", report)
        self.assertIn("B1. 2.2.: Pievienot nostiprināšanu. Piedāvātais formulējums: „Puses nostiprina īri zemesgrāmatā.”", report)
        for code in ("K1", "B1", "B2", "B3"):
            self.assertEqual(report.count(f"\n{code}. ") + report.startswith(code), 2, code)  # sadaļā + ieteikumos
        self.assertTrue(ctx.complete)

    def test_missing_ids_trigger_repair_and_never_lost(self):
        data = self.good()
        data["items"] = [i for i in data["items"] if i["id"] != "RENT_AMOUNT"]
        repair = {"items": [item("RENT_AMOUNT", "kritisks", atrasts=False, nosaukums="Īres summa", problema="Summa nav norādīta.")]}
        f = Fake(data, l1_repair=repair, l2={"items": []})
        report, info = ae.run_analysis(make_ctx(f))
        self.assertIn("Īres summa", report)
        self.assertIn("26/26" if len(ae.CHECK_IDS) == 26 else f"{len(ae.CHECK_IDS)}/{len(ae.CHECK_IDS)}", info)

    def test_unrepairable_missing_is_reported_and_not_cached(self):
        data = self.good()
        data["items"] = [i for i in data["items"] if i["id"] != "TERM"]
        f = Fake(data, l1_repair={"items": []}, l2={"items": []})
        ctx = make_ctx(f)
        _, info = ae.run_analysis(ctx)
        self.assertIn("TERM", info)
        self.assertFalse(ctx.complete)

    def test_fabricated_quote_is_dropped_but_finding_kept(self):
        data = self.good()
        for i in data["items"]:
            if i["id"] == "ENTRY":
                i["originalais_teksts"] = "šāda teksta līgumā nav"
        f = Fake(data, l2={"items": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertIn("B1. 6.1. (Apskate)", report)
        self.assertNotIn("šāda teksta", report)

    def test_money_criterion_law_violation_becomes_risk_factor(self):
        data = self.good()
        data["items"].append(item("RENT_CHANGE", "bridinajums", punkts="3.5", nosaukums="Maksas celšana",
                                  originalais_teksts="jebkurā laikā", problema="Vienpusēja maksas celšana."))
        f = Fake(data, l2={"items": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertIn("(Maksas celšana)", report.split("## 🟡")[0])

    def test_money_criterion_vague_clause_becomes_risk_factor(self):
        data = self.good()
        data["items"].append(item("UTILITIES", "apsverams", punkts="4.1", nosaukums="Komunālie",
                                  originalais_teksts="jebkurā laikā", problema="Neskaidrs sadalījums."))
        f = Fake(data, l2={"items": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertIn("(Komunālie)", report.split("## 🟡")[0])

    def test_missing_money_clause_is_risk_factor(self):
        data = self.good()
        for i in data["items"]:
            if i["id"] == "UTILITIES":
                i.update(riska_limenis="apsverams", atrasts=False, nosaukums="Komunālie",
                         problema="Nav minēti.", ko_mainit="Pievienot")
        f = Fake(data, l1_repair=data, l2={"items": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertIn("Komunālie: Nav minēti", report.split("## 🟡")[0])

    def test_floor_for_missing_mandatory_clause(self):
        data = self.good()
        for i in data["items"]:
            if i["id"] == "LAND_REGISTRY":
                i["riska_limenis"] = "apsverams"; i["atrasts"] = False; i["originalais_teksts"] = ""
        f = Fake(data, l1_repair=data, l2={"items": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertIn("## 🔴 Kritiski\nK", report)

    def test_multiple_findings_per_criterion_and_sweep(self):
        data = self.good()
        data["items"].append(item("ENTRY", "bridinajums", punkts="6.2", nosaukums="Apskate 2",
                                  originalais_teksts="Izīrētājs drīkst ienākt dzīvoklī", problema="Otrs ieejas pārkāpums."))
        f = Fake(data, l2={"items": []})
        f.sweep = {"papildus": [item("CITS", "apsverams", punkts="1", nosaukums="Adrese",
                                     originalais_teksts="dzīvokli Rīgā, Brīvības iela 1-5", problema="Nav stāva.")]}
        report, info = ae.run_analysis(make_ctx(f))
        self.assertIn("Otrs ieejas pārkāpums", report)
        self.assertIn("Nav stāva", report)
        self.assertIn("(+1)", info)

    def test_proofread_two_passes_numbered(self):
        calls = []
        def llm(parts, mt):
            calls.append(1)
            if len(calls) == 1:
                return json.dumps({"kludas": [{"punkts": "6.1", "kluda": "Kļūda B", "citats": "jebkurā laikā", "labojums": "jebkurā brīdī"}]}), (1, 1)
            return json.dumps({"kludas": [{"punkts": "1", "kluda": "Kļūda A", "citats": "izīrē dzīvokli", "labojums": "izīrē dzīvokli."}]}), (1, 1)
        ctx = ae.AnalysisContext(kind="txt", content=CONTRACT, lang_name="latviešu", labels=LABELS, llm=llm)
        text, ok = ae.run_proofread(ctx, {"h_neutral": "Neitrālās", "label_fix": "Labojums"})
        self.assertTrue(ok)
        self.assertEqual(len(calls), 2)
        self.assertIn("N1. 1. Kļūda A", text)
        self.assertIn("N2. 6.1. Kļūda B", text)

    def test_layer2_failure_keeps_result_but_not_cached(self):
        f = Fake(self.good(), l2=RuntimeError("boom"))
        ctx = make_ctx(f)
        report, info = ae.run_analysis(ctx)
        self.assertIn("K1.", report)
        self.assertFalse(ctx.complete)

    def test_order_is_deterministic_by_clause(self):
        a = self.good(); b = self.good(); b["items"] = list(reversed(b["items"]))
        r1, _ = ae.run_analysis(make_ctx(Fake(a, l2={"items": []})))
        r2, _ = ae.run_analysis(make_ctx(Fake(b, l2={"items": []})))
        self.assertEqual(r1, r2)

    def test_no_findings_skips_layer2(self):
        f = Fake({"kopsavilkums": "OK", "items": full_items({}), "papildus": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertEqual(f.calls, ["l1", "sweep"])
        self.assertIn("## 🔴 Kritiski\n–", report)


    def test_money_rules(self):
        data = self.good()
        data["skaitli"] = {"ires_summa_mened": 400, "depozits": 1200, "depozita_atmaksa_darba_dienas": 30,
                           "soda_likme_diena_proc": 0.5, "tuksi_naudas_lauki": False, "neierobezoti_maksajumi": False}
        ctx = make_ctx(Fake(data))
        report, _ = ae.run_analysis(ctx)
        self.assertIn("pārsniedz 2 mēnešu īres maksu", report)
        self.assertIn("atmaksas termiņš (30 darba dienas)", report)
        self.assertIn("0.5% dienā", report)

    def test_money_content_becomes_red(self):
        data = self.good()
        data["papildus"] = [item("CITS", "bridinajums", punkts="7.1", originalais_teksts="Izīrētājs drīkst ieturēt drošības naudu",
                                 problema="Neskaidrs.", ko_mainit="Precizēt")]
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        red = report.split("## 🟡")[0]
        self.assertIn("7.1", red)


class Cache(unittest.TestCase):
    def test_roundtrip_and_keys(self):
        with tempfile.TemporaryDirectory() as d:
            c = ae.AnalysisCache(os.path.join(d, "t.db"))
            h = ae.content_hash("Līgums  teksts\n")
            self.assertEqual(h, ae.content_hash("Līgums teksts"))      # atstarpes normalizētas
            self.assertIsNone(c.get(h, "v1", "m", "lv"))
            c.put(h, "v1", "m", "lv", {"analysis": "X"})
            self.assertEqual(c.get(h, "v1", "m", "lv")["analysis"], "X")
            for other in (("v2", "m", "lv"), ("v1", "m2", "lv"), ("v1", "m", "en")):
                self.assertIsNone(c.get(h, *other))


if __name__ == "__main__":
    unittest.main()
