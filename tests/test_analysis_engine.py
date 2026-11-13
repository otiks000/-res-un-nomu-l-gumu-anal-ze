import json, os, sys, tempfile, unittest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import analysis_engine as ae

CONTRACT = """1. Izīrētājs izīrē dzīvokli Rīgā, Brīvības iela 1-5.
2.2. Līgums netiek nostiprināts zemesgrāmatā.
4.3. Izīrētājs drīkst ieturēt drošības naudu pēc saviem ieskatiem.
6.1. Izīrētājs drīkst ienākt dzīvoklī jebkurā laikā."""
LABELS = {"h_overall": "Kopējais", "h_crit": "Kritiski", "h_warn": "Brīdinājumi", "h_money": "Nauda",
          "h_neutral": "Neitrālās", "h_recs": "Ieteikumi", "label_proposed": "Piedāvātais formulējums", "label_translation": "Tulkojums"}


def item(cid, kind="nav", **kw):
    base = {"id": cid, "atrasts": True, "veids": kind, "pazimes": {}, "panti": [], "punkts": "", "nosaukums": "",
            "originalais_teksts": "", "problema": "", "ko_mainit": "", "labojums": "", "labojums_tulkojums": ""}
    base.update(kw)
    return base


def full_items(over):
    return [over.get(cid, item(cid)) for cid in ae.CHECK_IDS]


class Fake:
    def __init__(self, l1, l1_repair=None):
        self.l1, self.l1_repair, self.calls = l1, l1_repair, []

    def llm(self, parts, max_tokens):
        text = str(parts)
        if "PAPILDU PĀRBAUDE" in text:
            self.calls.append("sweep")
            return json.dumps(getattr(self, "sweep", {"papildus": []})), (100, 10)
        if "IEPRIEKŠĒJĀ ATBILDĒ" in text:
            self.calls.append("repair")
            return json.dumps(self.l1_repair or self.l1), (100, 50)
        self.calls.append("l1")
        return json.dumps(self.l1), (100, 50)


def make_ctx(fake, kind="txt"):
    return ae.AnalysisContext(kind=kind, content=CONTRACT, lang_name="latviešu", labels=LABELS, llm=fake.llm,
                              valid_articles=lambda c: {8, 12, 28, 35} if c == "DTIL" else {1716})


def top(report):
    return report.split("## 🟡")[0]


class Engine(unittest.TestCase):
    def good(self):
        return {"kopsavilkums": "Neparakstīt.", "papildus": [], "items": full_items({
            "LAND_REGISTRY": item("LAND_REGISTRY", "likuma_parkapums", punkts="2.2", nosaukums="Zemesgrāmata",
                                  originalais_teksts="Līgums netiek nostiprināts zemesgrāmatā", problema="Nav nostiprināts.",
                                  panti=[{"likums": "DTIL", "nr": 28}, {"likums": "DTIL", "nr": 999}],
                                  ko_mainit="Pievienot nostiprināšanu", labojums="Puses nostiprina īri zemesgrāmatā."),
            "DISPUTES": item("DISPUTES", "neskaidrs_vai_truksts", punkts="9.1", nosaukums="Strīdi",
                             originalais_teksts="jebkurā laikā", problema="Jurisdikcija nav noteikta."),
            "DEPOSIT_USE": item("DEPOSIT_USE", "likuma_parkapums", punkts="4.3", nosaukums="Depozīts",
                                originalais_teksts="drošības naudu pēc saviem ieskatiem", problema="Patvaļīga ieturēšana.",
                                panti=[{"likums": "DTIL", "nr": 12}], ko_mainit="Precizēt"),
            "ENTRY": item("ENTRY", "citi_nosacijums", punkts="6.1", nosaukums="Apskate",
                          originalais_teksts="jebkurā laikā", problema="Izīrētājs drīkst ienākt jebkurā laikā."),
            "RENT_PAYMENT": item("RENT_PAYMENT", "naudas_nosacijums", punkts="1", nosaukums="Maksāšana",
                                 originalais_teksts="Izīrētājs izīrē dzīvokli Rīgā", problema="Īre tiek maksāta mēnesī."),
        })}

    def test_full_run_sections_by_python(self):
        f = Fake(self.good())
        ctx = make_ctx(f)
        report, info = ae.run_analysis(ctx)
        self.assertEqual(sorted(f.calls), ["l1", "l1", "l1", "sweep"])        # 3 kritēriju grupas + papildu pārbaude
        self.assertIn("K1. 4.3. (Depozīts)", top(report))                    # panta DTIL 12 pārkāpums = 🔴
        self.assertIn("[[DTIL 12]]", report)
        self.assertIn("B1. 2.2. (Zemesgrāmata)", report)                     # panta DTIL 28 pārkāpums = 🟡
        self.assertIn("[[DTIL 28]]", report)
        self.assertNotIn("999", report)                                      # neesošs pants atmests
        self.assertIn("B2. 9.1. (Strīdi)", report)
        self.assertIn("## 💶 Nauda\nM1. 1. (Maksāšana): Īre tiek maksāta mēnesī.", report)   # naudas fakts bez vērtējuma
        self.assertTrue(report.rstrip().split("## ⚪ Neitrālās\n")[1].startswith("N1. 6.1. (Apskate)"))
        self.assertEqual(ctx.neutral_count, 1)
        for code in ("K1", "B1", "B2"):
            self.assertEqual(report.count(f"\n{code}. "), 2, code)           # sadaļā + ieteikumos
        self.assertNotIn("\nM1. 1.: ", report)                               # 💶 un ⚪ nav ieteikumos
        self.assertTrue(ctx.complete)

    def test_missing_ids_trigger_repair_and_never_lost(self):
        data = self.good()
        data["items"] = [i for i in data["items"] if i["id"] != "RENT_AMOUNT"]
        repair = {"items": [item("RENT_AMOUNT", "neskaidrs_vai_truksts", atrasts=False, nosaukums="Īres summa", problema="Summa nav norādīta.")]}
        report, info = ae.run_analysis(make_ctx(Fake(data, l1_repair=repair)))
        self.assertIn("Īres maksas apmērs", top(report))      # punktam bez numura nosaukums ir fiksēts                             # trūkstoša naudas summa = 🔴
        self.assertIn(f"{len(ae.CHECK_IDS)}/{len(ae.CHECK_IDS)}", info)

    def test_unrepairable_missing_is_reported_and_not_cached(self):
        data = self.good()
        data["items"] = [i for i in data["items"] if i["id"] != "TERM"]
        ctx = make_ctx(Fake(data, l1_repair={"items": []}))
        _, info = ae.run_analysis(ctx)
        self.assertIn("TERM", info)
        self.assertFalse(ctx.complete)

    def test_fabricated_quote_is_dropped_but_finding_kept(self):
        data = self.good()
        for i in data["items"]:
            if i["id"] == "DISPUTES":
                i["originalais_teksts"] = "šāda teksta līgumā nav"
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertIn("(Strīdi)", report)
        self.assertNotIn("šāda teksta", report)

    def test_objective_money_flag_makes_it_red(self):
        data = self.good()
        data["items"].append(item("RENT_CHANGE", "naudas_nosacijums", punkts="3.5", nosaukums="Maksas celšana",
                                  originalais_teksts="jebkurā laikā", problema="Izīrētājs maina maksu.",
                                  pazimes={"vienpusejs": True}))
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertIn("(Maksas celšana)", top(report))

    def test_money_fact_without_flags_is_neutral(self):
        report, _ = ae.run_analysis(make_ctx(Fake(self.good())))
        self.assertNotIn("(Maksāšana)", top(report))
        self.assertIn("(Maksāšana)", report.split("## 💶")[1])

    def test_money_fact_not_duplicated_when_already_flagged(self):
        data = self.good()
        data["items"].append(item("DEPOSIT_AMOUNT", "naudas_nosacijums", punkts="4.3", nosaukums="Depozīta apmērs",
                                  originalais_teksts="drošības naudu pēc saviem ieskatiem", problema="Depozīta fakts."))
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertNotIn("Depozīta fakts", report)           # 4.3 jau ir 🔴
        self.assertIn("(Maksāšana)", report.split("## 💶")[1])  # citi naudas fakti paliek

    def test_law_violation_without_article_depends_on_money(self):
        data = self.good()
        data["items"].append(item("UTILITIES", "likuma_parkapums", punkts="4.1", nosaukums="Komunālie",
                                  originalais_teksts="jebkurā laikā", problema="Pārkāpums."))
        data["items"].append(item("OCCUPANTS", "likuma_parkapums", punkts="6.2", nosaukums="Iemitinātie",
                                  originalais_teksts="jebkurā laikā", problema="Pārkāpums."))
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertIn("(Komunālie)", top(report))
        self.assertIn("(Iemitinātie)", report.split("## 🟡")[1].split("## 💶")[0])

    def test_missing_money_clause_is_red(self):
        data = self.good()
        for i in data["items"]:
            if i["id"] == "UTILITIES":
                i.update(veids="neskaidrs_vai_truksts", atrasts=False, nosaukums="Komunālie", problema="Nav minēti.")
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertIn("Komunālie maksājumi: Nav minēti", top(report))

    def test_floor_for_missing_mandatory_clause(self):
        data = self.good()
        for i in data["items"]:
            if i["id"] == "LAND_REGISTRY":
                i.update(atrasts=False, originalais_teksts="", panti=[])
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertIn("## 🔴 Kritiski\nK", report)

    def test_multiple_findings_per_criterion_and_sweep(self):
        data = self.good()
        data["items"].append(item("ENTRY", "neskaidrs_vai_truksts", punkts="6.2", nosaukums="Apskate 2",
                                  originalais_teksts="Izīrētājs drīkst ienākt dzīvoklī", problema="Otrs ieejas jautājums."))
        f = Fake(data)
        f.sweep = {"papildus": [item("CITS", "neskaidrs_vai_truksts", punkts="1", nosaukums="Adrese",
                                     originalais_teksts="dzīvokli Rīgā, Brīvības iela 1-5", problema="Nav stāva.")]}
        report, info = ae.run_analysis(make_ctx(f))
        self.assertIn("Otrs ieejas jautājums", report)
        self.assertIn("Nav stāva", report)
        self.assertIn("(+1)", info)

    def test_numeric_money_rules(self):
        data = self.good()
        data["skaitli"] = {"ires_summa_mened": 400, "depozits": 1200, "depozita_atmaksa_darba_dienas": 30,
                           "soda_likme_diena_proc": 0.5, "tuksi_naudas_lauki": False, "neierobezoti_maksajumi": False}
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertIn("pārsniedz 2 mēnešu īres maksu", top(report))
        self.assertIn("atmaksas termiņš (30 darba dienas)", top(report))
        self.assertIn("0.5% dienā", top(report))

    def test_proofread_two_passes_numbered_and_merged(self):
        calls = []
        def llm(parts, mt):
            calls.append(1)
            if len(calls) == 1:
                return json.dumps({"kludas": [{"punkts": "6.1", "kluda": "Kļūda B", "citats": "jebkurā laikā", "labojums": "jebkurā brīdī"}]}), (1, 1)
            return json.dumps({"kludas": [{"punkts": "1", "kluda": "Kļūda A", "citats": "izīrē dzīvokli", "labojums": "izīrē dzīvokli."}]}), (1, 1)
        ctx = ae.AnalysisContext(kind="txt", content=CONTRACT, lang_name="latviešu", labels=LABELS, llm=llm)
        labels = {"h_neutral": "Neitrālās", "label_fix": "Labojums"}
        text, ok = ae.run_proofread(ctx, labels)
        self.assertTrue(ok)
        self.assertEqual(len(calls), 2)
        self.assertIn("## ⚪ Neitrālās", text)
        self.assertIn("N1. 1. Kļūda A", text)
        self.assertIn("N2. 6.1. Kļūda B", text)
        calls.clear()
        text, _ = ae.run_proofread(ctx, labels, start=3, with_head=False)    # turpina piezīmju numerāciju zem esošās ⚪ sadaļas
        self.assertNotIn("##", text)
        self.assertIn("N3. 1. Kļūda A", text)

    def test_dwelling_article_8_does_not_escalate(self):
        f = {"id": "PREMISES", "atrasts": True, "kind": "likuma_parkapums", "flags": dict.fromkeys(ae.FLAGS, False),
             "articles": [("DTIL", 8)], "force_red": False}
        self.assertEqual(ae.classify(f), "bridinajums")
        g = dict(f, id="DEPOSIT_USE", articles=[("DTIL", 12)])
        self.assertEqual(ae.classify(g), "kritisks")
        h = dict(f, id="ENTRY", articles=[("DTIL", 18)])         # naudas pants nekāpina ne-naudas kritēriju
        self.assertEqual(ae.classify(h), "bridinajums")

    def test_duplicates_for_same_clause_are_merged(self):
        def mk(cid, level, arts):
            return {"id": cid, "level": level, "punkts": "5.1", "quote": "Īrnieks veic kapitālo remontu", "articles": arts,
                    "problema": cid, "ko_mainit": "", "labojums": "", "tulkojums": "", "nosaukums": cid}
        out = ae.merge_duplicates([mk("REPAIRS_MAJOR", "kritisks", [("DTIL", 18)]), mk("FUND", "kritisks", []),
                                   mk("LIABILITY", "bridinajums", [("DTIL", 18)])])
        self.assertEqual(len(out), 1)
        other = mk("TERM", "bridinajums", [("DTIL", 9)])          # cits līmenis, cits kritērijs, nav kopīgu pantu -> paliek atsevišķi
        self.assertEqual(len(ae.merge_duplicates([mk("REPAIRS_MAJOR", "kritisks", [("DTIL", 18)]), other])), 2)

    def test_compare_ignores_model_titles_and_duplicates(self):
        a = "## 🔴 K\nK1. 3.2. (Nokavējuma procenti): X.\n> „q”\n\nK2. 3.2. (Līgumsodi): Y.\n\n## 🟡 B\nB1. Pagarināšana: Z.\n"
        b = "## 🔴 K\nK1. 3.2. (Līgumsoda griesti): X.\n\n## 🟡 B\nB1. Pagarināšana: Z.\n"
        c = ae.compare_reports([a, b])
        self.assertEqual(len(c["rows"]), 2)
        self.assertEqual(c["stability"], 1.0)

    def test_order_is_deterministic_by_clause(self):
        a = self.good(); b = self.good(); b["items"] = list(reversed(b["items"]))
        r1, _ = ae.run_analysis(make_ctx(Fake(a)))
        r2, _ = ae.run_analysis(make_ctx(Fake(b)))
        self.assertEqual(r1, r2)

    def test_no_findings(self):
        f = Fake({"kopsavilkums": "OK", "items": full_items({}), "papildus": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertEqual(sorted(f.calls), ["l1", "l1", "l1", "sweep"])
        self.assertIn("## 🔴 Kritiski\n–", report)
        self.assertNotIn("## ⚪", report)


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


class CompareTests(unittest.TestCase):
    A = "## 🔴 K\nK1. 4.3. (Depozīts): slikti [[DTIL 12]]\n> „teksts”\n\n## 🟡 B\nB1. 6.1. (Ieeja): ok\n\n## 🟢 R\nK1. 4.3.: Mainīt. Piedāvātais: „Labāk”\n"
    B = "## 🔴 K\nK1. 4.3. (Depozīts): slikti\n\n## 🟡 B\n–\n"

    def test_parse(self):
        items = ae.parse_report(self.A)
        self.assertEqual([i["code"] for i in items], ["K1", "B1"])
        self.assertEqual(items[0]["fix"], "Labāk")
        self.assertEqual(items[0]["quote"], "„teksts”")

    def test_compare(self):
        r = ae.compare_reports([self.A, self.B, self.A])
        st = {row["ref"]: row["status"] for row in r["rows"]}
        self.assertEqual(st["4.3. (Depozīts)"], "stable")
        self.assertEqual(st["6.1. (Ieeja)"], "partial")
        self.assertAlmostEqual(r["stability"], 0.5)
        self.assertEqual(r["counts"][1]["bridinajums"], 0)

    def test_level_change(self):
        b = self.A.replace("## 🔴 K\nK1.", "## 🔴 K\nX9.").replace("## 🟡 B\nB1. 6.1.", "## 🟡 B\nB1. 4.3. (Depozīts): z\nB2. 6.1.")
        r = ae.compare_reports([self.A, b])
        self.assertIn("level", {row["status"] for row in r["rows"]})


if __name__ == "__main__":
    unittest.main()
