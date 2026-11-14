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
        self.assertEqual(sorted(f.calls), ["l1", "l1", "l1"])        # tikai 3 kritēriju grupas, bez papildu pārbaudes
        self.assertIn("K2. 4.3. (Depozīts)", top(report))                    # panta DTIL 12 pārkāpums = 🔴
        self.assertIn("[[DTIL 12]]", report)
        self.assertIn("K1. 2.2. (Zemesgrāmata)", top(report))                # zemesgrāmatas problēma = 🔴 (fiksēts saraksts)
        self.assertIn("[[DTIL 28]]", report)
        self.assertNotIn("999", report)                                      # neesošs pants atmests
        self.assertIn("B1. 6.1. (Strīdi)", report)                        # punkta numuru nosaka citāts līgumā, nevis modelis
        self.assertIn("## 💶 Nauda\nM1. 1. (Maksāšana): Īre tiek maksāta mēnesī.", report)   # naudas fakts bez vērtējuma
        self.assertTrue(report.rstrip().split("## ⚪ Neitrālās\n")[1].startswith("N1. 6.1. (Apskate)"))
        self.assertEqual(ctx.neutral_count, 1)
        for code in ("K1", "K2", "B1"):
            self.assertEqual(report.count(f"\n{code}. "), 2, code)           # sadaļā + ieteikumos
        self.assertNotIn("\nM1. 1.: ", report)                               # 💶 un ⚪ nav ieteikumos
        self.assertTrue(ctx.complete)

    def test_missing_ids_trigger_repair_and_never_lost(self):
        data = self.good()
        data["items"] = [i for i in data["items"] if i["id"] != "RENT_AMOUNT"]
        repair = {"items": [item("RENT_AMOUNT", "neskaidrs_vai_truksts", atrasts=False, nosaukums="Īres summa", problema="Summa nav norādīta.")]}
        ctx = make_ctx(Fake(data, l1_repair=repair))
        report, info = ae.run_analysis(ctx)
        self.assertIn("Īres maksas apmērs", top(report))      # punktam bez numura nosaukums ir fiksēts; trūkstoša summa = 🔴
        n = len(ctx.check_ids)
        self.assertEqual(n, len([c for c in ae.CHECK_IDS if ae.applies(c, "dz")]))
        self.assertIn(f"{n}/{n}", info)

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
        data["items"].append(item("OCCUPANTS", "likuma_parkapums", punkts="1", nosaukums="Iemitinātie",
                                  originalais_teksts="Izīrētājs izīrē dzīvokli Rīgā", problema="Pārkāpums."))
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
            if i["id"] == "TERM":
                i.update(veids="neskaidrs_vai_truksts", atrasts=False, originalais_teksts="", panti=[], problema="Nav termiņa.")
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertIn("## 🔴 Kritiski\nK", report)

    def test_multiple_findings_per_criterion_and_extras(self):
        data = self.good()
        data["items"].append(item("ENTRY", "neskaidrs_vai_truksts", punkts="6.2", nosaukums="Apskate 2",
                                  originalais_teksts="Izīrētājs drīkst ienākt dzīvoklī", problema="Otrs ieejas jautājums."))
        data["papildus"] = [item("CITS", "neskaidrs_vai_truksts", punkts="1", nosaukums="Adrese",
                                 originalais_teksts="dzīvokli Rīgā, Brīvības iela 1-5", problema="Nav stāva.")]
        f = Fake(data)
        report, info = ae.run_analysis(make_ctx(f))
        self.assertIn("Otrs ieejas jautājums", report)
        self.assertIn("Nav stāva", report)                    # papildu atradumi nāk no grupām (papildus), atsevišķa izsaukuma nav
        self.assertNotIn("sweep", f.calls)

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
        h = dict(f, id="RESTRICTIONS", articles=[("DTIL", 18)])  # naudas pants nekāpina ne-naudas kritēriju
        self.assertEqual(ae.classify(h), "bridinajums")

    def test_duplicates_for_same_clause_are_merged(self):
        def mk(cid, level, arts, punkts="5.1", problema=None):
            return {"id": cid, "level": level, "punkts": punkts, "quote": "q", "articles": arts,
                    "problema": problema or cid, "ko_mainit": "", "labojums": "", "tulkojums": "", "nosaukums": cid}
        out = ae.merge_duplicates([mk("REPAIRS_MAJOR", "kritisks", [("DTIL", 18)]), mk("FUND", "kritisks", []),
                                   mk("RENT_CHANGE", "kritisks", [("DTIL", 10)])])
        self.assertEqual(len(out), 1)                              # tas pats punkts + tas pats līmenis = viens ieraksts
        self.assertIn("FUND", out[0]["problema"])                  # problēmu teksti saglabāti
        self.assertEqual(len(out[0]["articles"]), 2)
        diff = ae.merge_duplicates([mk("REPAIRS_MAJOR", "kritisks", []), mk("TERM", "bridinajums", [])])
        self.assertEqual(len(diff), 2)                             # cits līmenis paliek atsevišķi
        other = ae.merge_duplicates([mk("REPAIRS_MAJOR", "kritisks", []), mk("FUND", "kritisks", [], punkts="6.1")])
        self.assertEqual(len(other), 2)                            # cits punkts paliek atsevišķi
        nonum = ae.merge_duplicates([mk("ENTRY", "bridinajums", [], punkts=""), mk("ENTRY", "bridinajums", [], punkts="")])
        self.assertEqual(len(nonum), 1)                            # bez numura apvieno pēc kritērija

    def test_merged_findings_keep_every_recommendation(self):
        def mk(cid, ko, fix):
            return {"id": cid, "level": "kritisks", "punkts": "5.1", "quote": "q", "articles": [], "problema": cid,
                    "ko_mainit": ko, "labojums": fix, "tulkojums": "", "nosaukums": cid}
        out = ae.merge_duplicates([mk("REPAIRS_MAJOR", "Svītrot remontu", "Remontu veic izīrētājs."),
                                   mk("FUND", "Svītrot fondu", "Fondu maksā izīrētājs.")])
        self.assertEqual(len(out), 1)
        f = dict(out[0], atrasts=True)
        report = ae.render_report([dict(out[0], level="kritisks")], "x", LABELS)
        rec = report.split("## 🟢")[1]
        self.assertIn("Svītrot remontu; Svītrot fondu", rec)
        self.assertIn("„Remontu veic izīrētājs.” „Fondu maksā izīrētājs.”", rec)

    def test_compare_ignores_model_titles_and_duplicates(self):
        a = "## 🔴 K\nK1. 3.2. (Nokavējuma procenti): X.\n> „q”\n\nK2. 3.2. (Līgumsodi): Y.\n\n## 🟡 B\nB1. Pagarināšana: Z.\n"
        b = "## 🔴 K\nK1. 3.2. (Līgumsoda griesti): X.\n\n## 🟡 B\nB1. Pagarināšana: Z.\n"
        c = ae.compare_reports([a, b])
        self.assertEqual(len(c["rows"]), 2)
        self.assertEqual(c["stability"], 1.0)

    def test_contract_type_detection(self):
        self.assertEqual(ae.detect_contract_type("Dzīvojamo telpu īres līgums. Īrnieks un Izīrētājs. Dzīvoklis."), "dz")
        self.assertEqual(ae.detect_contract_type("Nedzīvojamo telpu nomas līgums. Nomnieks, Iznomātājs, birojs."), "noma")
        self.assertEqual(ae.detect_contract_type("Līgums bez nekādām norādēm."), "")

    def test_fixed_level_table(self):
        base = {"atrasts": True, "flags": dict.fromkeys(ae.FLAGS, False), "articles": [], "force_red": False}
        c = lambda cid, kind, **kw: ae.classify(dict(base, id=cid, kind=kind, **kw))
        self.assertEqual(c("RESTRICTIONS", "likuma_parkapums", articles=[("DTIL", 18)]), "bridinajums")   # panti līmeni nemaina
        for kind in ("likuma_parkapums", "neskaidrs_vai_truksts"):                                    # likuma kritērijs: krāsa nav atkarīga no modeļa veida
            self.assertEqual(c("TERM", kind), "kritisks")
            self.assertEqual(c("ENTRY", kind), "kritisks")
            self.assertEqual(c("LANDLORD_NOTICE", kind), "kritisks")
        for kind in ("likuma_parkapums", "neskaidrs_vai_truksts"):                                    # zemesgrāmatas aizliegums un klusēšana = 🔴
            self.assertEqual(c("LAND_REGISTRY", kind), "kritisks")
        self.assertEqual(c("LAND_REGISTRY", "neskaidrs_vai_truksts", atrasts=False), "kritisks")     # klusēšana arī = 🔴
        flag = lambda **kw: dict(base["flags"], **kw)
        self.assertEqual(c("BLANKS", "neskaidrs_vai_truksts", flags=flag(tukss=True)), "kritisks")        # tukša summa arī šajā kritērijā
        self.assertEqual(c("BLANKS", "neskaidrs_vai_truksts"), "bridinajums")
        self.assertEqual(c("CITS", "neskaidrs_vai_truksts", flags=flag(bez_griestiem=True)), "kritisks")  # papildu atradums ar naudas pazīmi
        self.assertEqual(c("CITS", "neskaidrs_vai_truksts"), "bridinajums")
        self.assertEqual(c("CITS", "likuma_parkapums", articles=[("DTIL", 28)]), "kritisks")         # papildu atradums ar derīgu pantu
        self.assertEqual(c("CITS", "likuma_parkapums"), "bridinajums")                              # bez panta (vai pants atmests) paliek 🟡
        noma = lambda cid, kind: ae.classify(dict(base, id=cid, kind=kind), "noma")
        self.assertEqual(noma("TERM", "likuma_parkapums"), "bridinajums")                           # nenoteikts termiņš nomā ir atļauts
        self.assertEqual(noma("LANDLORD_NOTICE", "neskaidrs_vai_truksts"), "kritisks")
        self.assertEqual(c("RENT_CHANGE", "neskaidrs_vai_truksts"), "kritisks")                     # naudas kritērijs = 🔴
        self.assertEqual(c("PARTIES", "neskaidrs_vai_truksts"), "bridinajums")
        self.assertEqual(c("TERM", "neskaidrs_vai_truksts", flags=dict(base["flags"], tukss=True)), "kritisks")  # tukšs termiņš
        self.assertEqual(c("RENT_AMOUNT", "naudas_nosacijums", flags=dict(base["flags"], tukss=True)), "kritisks")
        self.assertEqual(c("RENT_AMOUNT", "naudas_nosacijums"), "nauda")
        self.assertEqual(c("TERM", "neskaidrs_vai_truksts", atrasts=False), "kritisks")             # trūkst obligāta tēma
        self.assertEqual(c("PARTIES", "neskaidrs_vai_truksts", atrasts=False), "bridinajums")
        self.assertEqual(c("PENALTY_CAP", "neskaidrs_vai_truksts", atrasts=False), "nav")           # trūkums netiek atzīmēts

    def test_clause_number_comes_from_contract_quote(self):
        ctx = make_ctx(Fake({}))
        f = {"quote": "Izīrētājs drīkst ienākt dzīvoklī jebkurā laikā", "punkts": "99"}
        ae.assign_clauses(ctx, [f])
        self.assertEqual(f["punkts"], "6.1")

    def test_no_deposit_means_no_deposit_findings(self):
        data = self.good()
        for i in data["items"]:
            if i["id"] == "DEPOSIT_AMOUNT":
                i.update(veids="nav", atrasts=False)
            if i["id"] == "DEPOSIT_REFUND":
                i.update(veids="neskaidrs_vai_truksts", atrasts=False, problema="Nav atmaksas.")
        report, _ = ae.run_analysis(make_ctx(Fake(data)))
        self.assertNotIn("Nav atmaksas", report)

    def test_commercial_lease_uses_lease_criteria_and_no_dtil(self):
        content = "Nedzīvojamo telpu nomas līgums.\n1. Nomnieks maksā nomas maksu.\n2. Nomas maksa ir _____ EUR."
        ctx = ae.AnalysisContext(kind="txt", content=content, lang_name="latviešu", labels=LABELS, llm=None,
                                 valid_articles=lambda c: {8, 12} if c == "DTIL" else {2151})
        prompts = []
        def llm(parts, mt):
            prompts.append(str(parts))
            items = [item("RENT_AMOUNT", "naudas_nosacijums", punkts="2", nosaukums="Maksa", originalais_teksts="Nomas maksa ir _____ EUR",
                          problema="Summa nav ierakstīta.", pazimes={"tukss": True}, panti=[{"likums": "DTIL", "nr": 12}])]
            return json.dumps({"lig_veids": "telpu_noma", "kopsavilkums": "x", "items": items, "papildus": []}), (1, 1)
        ctx.llm = llm
        report, info = ae.run_analysis(ctx)
        joined = "".join(prompts)
        self.assertEqual(ctx.contract_type, "noma")
        self.assertIn("telpu noma", info)
        self.assertNotIn('id="OCCUPANTS"', joined)                       # dzīvojamās īres kritērijs nomai netiek prasīts
        self.assertIn('id="VAT"', joined)                                # nomas kritērijs tiek prasīts
        self.assertNotIn("DTIL", joined.split("DOTIE KRITĒRIJI")[1].split("LĪGUMA TEKSTS")[0])
        self.assertIn("K1. 2. (Maksa)", top(report))                     # tukša cena = 🔴
        self.assertNotIn("[[DTIL", report)                               # nomai DTIL pants tiek atmests

    def test_unknown_type_is_decided_by_model(self):
        ctx = ae.AnalysisContext(kind="pdf", content=b"x", lang_name="latviešu", labels=LABELS, llm=None)
        def llm(parts, mt):
            return json.dumps({"lig_veids": "telpu_noma", "kopsavilkums": "x", "items": [], "papildus": []}), (1, 1)
        ctx.llm = llm
        ae.run_analysis(ctx)
        self.assertEqual(ctx.contract_type, "noma")
        self.assertNotIn("OCCUPANTS", ctx.check_ids)

    def test_order_is_deterministic_by_clause(self):
        a = self.good(); b = self.good(); b["items"] = list(reversed(b["items"]))
        r1, _ = ae.run_analysis(make_ctx(Fake(a)))
        r2, _ = ae.run_analysis(make_ctx(Fake(b)))
        self.assertEqual(r1, r2)

    def test_no_findings(self):
        f = Fake({"kopsavilkums": "OK", "items": full_items({}), "papildus": []})
        report, _ = ae.run_analysis(make_ctx(f))
        self.assertEqual(sorted(f.calls), ["l1", "l1", "l1"])
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
