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
            "DEPOSIT_USE": item("DEPOSIT_USE", "kritisks", punkts="4.3", nosaukums="Depozīts",
                                originalais_teksts="drošības naudu pēc saviem ieskatiem", problema="Patvaļīga ieturēšana.", ko_mainit="Precizēt"),
            "ENTRY": item("ENTRY", "apsverams", punkts="6.1", nosaukums="Apskate",
                          originalais_teksts="jebkurā laikā", problema="Nav brīdinājuma.", ko_mainit="Pievienot 24h"),
        })}

    def test_full_run_and_layer2(self):
        l2 = {"kopsavilkums": "Pēc pārbaudes.", "items": [
            {"id": "F2", "verdikts": "apstiprinats", "panti": [{"likums": "DTIL", "nr": 28}, {"likums": "DTIL", "nr": 999}]},
            {"id": "F0", "verdikts": "pazeminat_bridinajums", "panti": [{"likums": "DTIL", "nr": 12}]}]}
        f = Fake(self.good(), l2=l2)
        ctx = make_ctx(f)
        report, info = ae.run_analysis(ctx)
        self.assertEqual(f.calls, ["l1", "l2"])
        self.assertIn("K1. 2.2. (Zemesgrāmata)", report)
        self.assertIn("[[DTIL 28]]", report)
        self.assertNotIn("999", report)                       # neesošs pants atmests
        self.assertIn("B1. 4.3. (Depozīts)", report)          # pazemināts, bet nav izdzēsts
        self.assertIn("A1. 6.1. (Apskate)", report)
        self.assertIn("Pēc pārbaudes.", report)
        self.assertIn("K1. 2.2.: Pievienot nostiprināšanu. Piedāvātais formulējums: „Puses nostiprina īri zemesgrāmatā.”", report)
        for code in ("K1", "B1", "A1"):
            self.assertEqual(report.count(f"\n{code}. ") + report.startswith(code), 2, code)  # sadaļā + ieteikumos
        self.assertTrue(ctx.complete)

    def test_missing_ids_trigger_repair_and_never_lost(self):
        data = self.good()
        data["items"] = [i for i in data["items"] if i["id"] != "RENT_AMOUNT"]
        repair = {"items": [item("RENT_AMOUNT", "bridinajums", atrasts=False, nosaukums="Īres summa", problema="Summa nav norādīta.")]}
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
        self.assertIn("A1. 6.1. (Apskate)", report)
        self.assertNotIn("šāda teksta", report)

    def test_floor_for_missing_mandatory_clause(self):
        data = self.good()
        for i in data["items"]:
            if i["id"] == "LAND_REGISTRY":
                i["riska_limenis"] = "apsverams"; i["atrasts"] = False; i["originalais_teksts"] = ""
        f = Fake(data, l1_repair=data, l2={"items": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertIn("## 🔴 Kritiski\nK", report)

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
        self.assertEqual(f.calls, ["l1"])
        self.assertIn("## 🔴 Kritiski\n–", report)


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
