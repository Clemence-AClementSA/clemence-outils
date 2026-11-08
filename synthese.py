"""Synthèse factures A. Clément SA – PDF corporate.
Usage : python3 synthese.py data.json OUT_DIR TEMPLATE_DIR
data.json :
{ "type": "direction" | "cp", "date": "AAAA-MM-JJ", "destinataire": "Samuel Page",
  "email": "samuel.page@a-clement.ch",
  "factures": [ {"board": "societe"|"chantier", "categorie": "Société"|"EG/ET"|"DT"|"Interne", "remarque": "...",
                 "statut": "...", "fournisseur": "...", "objet": "...", "affaire": "026005 – Charmey",
                 "cfc": "258", "no_facture": "...", "date_facture": "AAAA-MM-JJ", "recue_le": "AAAA-MM-JJ",
                 "montant_ht": 0.0, "montant_ttc": 0.0, "escompte": "2% 10 j", "fin_escompte": "AAAA-MM-JJ",
                 "echeance": "AAAA-MM-JJ", "a_valider_avant": "AAAA-MM-JJ", "cp": "Samuel Page",
                 "url": "https://...", "maj": "AAAA-MM-JJ"} ],
  "anomalies": ["texte", ...],
  "test": false }   # true pendant la période de test : bandeau TEST + « TEST » dans l'objet du nom de fichier
Produit OUT_DIR/ACFIN_Rapport_<prénom nom>_Synthèse factures[ direction][ TEST]_<AAAA.MM.JJ>.pdf (+ .docx)
  <prénom nom> = partie de l'adresse e-mail avant @a-clement.ch, point remplacé par un espace (ex. « alain clement »).
  Nommage conforme ASQ_ACSA_026.02 (code société, type, contrepartie, objet, date en dernier, aucun « __ »).
  Le flux Power Automate « Factures – envoi des synthèses » reconstruit l'adresse depuis ce champ.
"""
import sys, json, os, subprocess, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from aclement import *
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]


def d(s):
    return dt.date.fromisoformat(s) if s else None


def fd(s):
    x = d(s)
    return x.strftime("%d.%m.%Y") if x else "–"


def chf(v):
    if v in (None, ""): return "–"
    s = f"{float(v):,.2f}".replace(",", "'")
    return s


def payer_avant(f):
    return f.get("fin_escompte") or f.get("echeance")


def add_link(cell, text, url):
    p = cell.paragraphs[0]
    if not url:
        r = p.add_run(text); return
    part = p.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), rid)
    r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
    f = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"): f.set(qn(a), "Roboto")
    c = OxmlElement("w:color"); c.set(qn("w:val"), "C9AA8A")
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single")
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "17")
    b = OxmlElement("w:b")
    for e in (f, b, c, sz, u): rpr.append(e)
    t = OxmlElement("w:t"); t.text = text
    r.append(rpr); r.append(t); h.append(r); p._p.append(h)


def tbl(doc, headers, rows, widths, links=None):
    """links: liste d'URL (une par ligne) placée dans la dernière colonne."""
    t = table(doc, headers, [r + [""] for r in rows] if links is not None else rows, widths, size=8.5)
    if links is not None:
        for i, url in enumerate(links):
            add_link(t.rows[i + 1].cells[-1], "PDF", url)
    return t


def empty(doc, txt="Aucune facture."):
    para(doc, txt, size=10, color="7F7F7F", after=6)


def test_banner(doc, data):
    if data.get("test"):
        callout(doc, "**PÉRIODE DE TEST** – cette synthèse est produite en mode test (jusqu'au "
                     + data.get("test_fin", "19.11.2026") + "). Les données sont réelles : merci de traiter les factures normalement "
                     "et de signaler toute erreur ou incompréhension à Alain.", fill="FBE9E7")


def build(data, out_dir, tpl_dir):
    today = d(data["date"])
    jour = JOURS[today.weekday()]
    doc = word_from_template(os.path.join(tpl_dir, "Modèles - Notes internes.docx"))
    F = data.get("factures", [])
    soon = today + dt.timedelta(days=7)

    if data["type"] == "direction":
        a_payer = [f for f in F if f["board"] == "societe" and f.get("statut") not in ("Payé",)] + \
                  [f for f in F if f["board"] == "chantier" and f["categorie"] == "EG/ET" and f.get("statut") == "Validé – bon créé"]
        attente = [f for f in F if f["board"] == "chantier" and f["categorie"] == "EG/ET" and f.get("statut") in ("À contrôler", "Urgent", "À clarifier")]
        a_payer.sort(key=lambda f: payer_avant(f) or "9999")
        attente.sort(key=lambda f: f.get("a_valider_avant") or "9999")
        tot = sum(float(f.get("montant_ttc") or 0) for f in a_payer)
        tot_att = sum(float(f.get("montant_ttc") or 0) for f in attente)
        n7 = [f for f in a_payer if payer_avant(f) and d(payer_avant(f)) <= soon]
        title(doc, "Synthèse factures – paiements",
              f"Édition du {jour} {today.strftime('%d.%m.%Y')} – confidentiel, direction")
        test_banner(doc, data)
        info_block(doc, [("Destinataire :", data.get("destinataire", "Alain Clément")), ("Date :", today.strftime("%d.%m.%Y")),
                         ("Périmètre :", "Société, EG, ET"), ("Source :", "Monday 01.1 et 01.2")])
        callout(doc, f"**OK à payer : {len(a_payer)} facture(s) – CHF {chf(tot)}**, dont {len(n7)} à payer d'ici le {soon.strftime('%d.%m.%Y')}.\n"
                     f"**En attente de validation (EG / ET) : {len(attente)} facture(s) – CHF {chf(tot_att)}.**")

        section(doc, "1. À payer – OK pour paiement")
        if a_payer:
            rows, links = [], []
            for f in a_payer:
                pa = payer_avant(f); late = pa and d(pa) < today
                rows.append([("**" + fd(pa) + "**") if (pa and d(pa) <= soon) else fd(pa),
                             f.get("fournisseur", ""), (f.get("affaire") or f.get("objet") or "")[:60],
                             f.get("no_facture", ""), chf(f.get("montant_ttc")),
                             f.get("escompte") or "–",
                             ("Société" if f["board"] == "societe" else "EG/ET validée") + (" – RETARD" if late else "")])
                links.append(f.get("url"))
            tbl(doc, ["Payer avant", "Fournisseur", "Objet / affaire", "N° facture", "TTC CHF", "Escompte", "Type", "Fichier"],
                rows, [2.2, 2.5, 3.2, 1.8, 1.8, 2.0, 2.0, 1.3], links)
        else:
            empty(doc)

        section(doc, "2. EG / ET – en attente de validation du chef de projet")
        if attente:
            rows, links = [], []
            for f in attente:
                av = f.get("a_valider_avant"); late = av and d(av) < today
                rows.append([("**" + fd(av) + " ⚠**") if late else fd(av), f.get("cp", ""), (f.get("affaire") or "")[:40],
                             f.get("fournisseur", ""), chf(f.get("montant_ttc")), f.get("statut", "")])
                links.append(f.get("url"))
            tbl(doc, ["À valider avant", "Chef de projet", "Affaire", "Fournisseur", "TTC CHF", "Statut", "Fichier"],
                rows, [2.2, 2.4, 3.7, 3.1, 1.9, 2.2, 1.3], links)
        else:
            empty(doc)

        clar = [f for f in F if f["board"] == "chantier" and f.get("statut") == "À clarifier"]
        section(doc, "3. À clarifier – factures bloquées (toutes affaires)")
        if clar:
            rows, links = [], []
            for f in clar:
                rows.append([f.get("cp", ""), (f.get("affaire") or "")[:34], f.get("fournisseur", ""), f.get("no_facture", ""),
                             chf(f.get("montant_ttc")), (f.get("remarque") or "–")[:120]])
                links.append(f.get("url"))
            tbl(doc, ["Chef de projet", "Affaire", "Fournisseur", "N° facture", "TTC CHF", "Remarque", "Fichier"],
                rows, [2.3, 2.9, 2.6, 1.8, 1.8, 4.1, 1.3], links)
        else:
            empty(doc, "Aucune facture à clarifier.")

        section(doc, "4. À trier / points d'attention")
        if data.get("anomalies"):
            for a in data["anomalies"]: bullet(doc, a, size=10)
        else:
            empty(doc, "Aucun point d'attention.")

        if today.weekday() == 0 or data.get("vue_semaine"):
            section(doc, "5. Vue de la semaine")
            week0 = today - dt.timedelta(days=7)
            recues = [f for f in F if f.get("recue_le") and d(f["recue_le"]) >= week0]
            retard = [f for f in F if f["board"] == "chantier" and f.get("statut") in ("À contrôler", "Urgent", "À clarifier")
                      and f.get("a_valider_avant") and d(f["a_valider_avant"]) < today]
            dt_att = [f for f in F if f["board"] == "chantier" and ((f["categorie"] == "DT" and f.get("statut") == "Validé – bon créé") or (f["categorie"] == "Interne" and f.get("statut") not in ("Ordre paiement MO envoyé", "Payé")))]
            dt_env = [f for f in F if f["board"] == "chantier" and f["categorie"] in ("DT", "Interne") and f.get("statut") == "Ordre paiement MO envoyé"
                      and f.get("maj") and d(f["maj"]) >= week0]
            def s(lst): return f"{len(lst)} – CHF {chf(sum(float(x.get('montant_ttc') or 0) for x in lst))}"
            table(doc, ["Indicateur", "Nombre – montant TTC"], [
                ["Factures reçues (7 derniers jours) – société", s([f for f in recues if f['board'] == 'societe'])],
                ["Factures reçues – chantiers EG / ET", s([f for f in recues if f.get('categorie') == 'EG/ET'])],
                ["Factures reçues – chantiers DT / mandat (payées par le MO)", s([f for f in recues if f.get('categorie') == 'DT'])],
                ["Factures internes A. Clément SA (honoraires) reçues", s([f for f in recues if f.get('categorie') == 'Interne'])],
                ["Factures à clarifier (toutes affaires)", s(clar)],
                ["Validations en retard (tous chefs de projet)", s(retard)],
                ["À mettre sur l'ordre de paiement MO du vendredi (DT et factures A. Clément SA)", s(dt_att)],
                ["DT et internes transmises aux maîtres d'ouvrage cette semaine", s(dt_env)],
            ], [11.0, 6.0], size=9, first_col_bold=True)
            if retard:
                rows = [[fd(f.get("a_valider_avant")), f.get("cp", ""), (f.get("affaire") or "")[:40], f.get("fournisseur", ""), chf(f.get("montant_ttc"))] for f in retard]
                subsection(doc, "Détail des validations en retard")
                table(doc, ["Échue le", "Chef de projet", "Affaire", "Fournisseur", "TTC CHF"], rows, [2.4, 3.0, 5.0, 4.4, 2.2], size=8.5)
    else:
        mine = [f for f in F if f["board"] == "chantier"]
        internes = [f for f in mine if f["categorie"] == "Interne"]
        ctrl = [f for f in mine if f["categorie"] != "Interne"]
        vend = today + dt.timedelta(days=(4 - today.weekday()) % 7)
        urg = [f for f in ctrl if f.get("statut") in ("À contrôler", "Urgent", "À clarifier") and
               (f.get("statut") == "Urgent" or (f.get("a_valider_avant") and d(f["a_valider_avant"]) <= today + dt.timedelta(days=2)))]
        OPEN = ("À contrôler", "Urgent")
        mo = [f for f in mine if f["categorie"] in ("DT", "Interne") and f.get("statut") in OPEN + ("Validé – bon créé",)]
        eg = [f for f in mine if f["categorie"] == "EG/ET" and f.get("statut") in OPEN]
        clar = [f for f in mine if f.get("statut") == "À clarifier"]
        recent = [f for f in mine if f.get("statut") in ("Validé – bon créé", "Ordre paiement MO envoyé", "Payé") and f.get("maj") and d(f["maj"]) >= today - dt.timedelta(days=7)]
        for l in (urg, mo, eg, clar): l.sort(key=lambda f: f.get("a_valider_avant") or "9999")
        mo_bon = [f for f in mo if f.get("statut") in OPEN]
        title(doc, "Synthèse factures – à contrôler et valider",
              f"Édition du {jour} {today.strftime('%d.%m.%Y')} – {data.get('destinataire', '')}")
        test_banner(doc, data)
        info_block(doc, [("Destinataire :", data.get("destinataire", "")), ("Date :", today.strftime("%d.%m.%Y")),
                         ("Board :", "Monday 01.2 Factures chantiers"), ("Procédure :", "ASQ_ACSA_026.10")])
        callout(doc, f"**{len(urg)} urgente(s) ou en retard – {len(mo)} pour l'ordre de paiement au MO du vendredi {vend.strftime('%d.%m.%Y')} (dont {len(mo_bon)} bon(s) de paiement à créer) – {len(eg)} bon(s) de paiement interne(s) EG / ET – {len(clar)} à clarifier.**\n"
                     "Pour chaque facture : contrôler → saisir dans Delta → créer et classer le bon de paiement → passer la tâche Programme chantiers à « Terminé » (ou la ligne 01.2 à « Validé – bon créé »). "
                     "Le vendredi : établir l'ordre de paiement dans Delta (il reprend tous les bons pas encore transmis) et l'envoyer au maître d'ouvrage.")

        def block(lst, n, ttl, afaire=False):
            section(doc, f"{n}. {ttl}")
            if not lst: empty(doc); return
            rows, links, tot = [], [], 0.0
            for f in lst:
                av = f.get("a_valider_avant"); late = av and d(av) < today and f.get("statut") in OPEN
                st = f.get("statut", "")
                if afaire:
                    st = "Bon créé → ordre du vendredi" if st == "Validé – bon créé" else ("Bon à créer (urgent)" if st == "Urgent" else "Bon à créer")
                typ = "A. Clément SA" if f["categorie"] == "Interne" else f["categorie"]
                rows.append([("**" + fd(av) + (" ⚠" if late else "") + "**") if (late or f.get("statut") == "Urgent") else fd(av),
                             (f.get("affaire") or "")[:34], f.get("cfc") or "–", f.get("fournisseur", ""), f.get("no_facture", ""),
                             chf(f.get("montant_ttc")), typ, st])
                links.append(f.get("url")); tot += float(f.get("montant_ttc") or 0)
            tbl(doc, ["À valider avant", "Affaire", "CFC", "Fournisseur", "N° facture", "TTC CHF", "Type", "À faire" if afaire else "Statut", "Fichier"],
                rows, [2.0, 2.6, 1.1, 2.4, 1.9, 1.7, 1.6, 2.3, 1.4], links)
            para(doc, f"Total : {len(lst)} facture(s) – CHF {chf(tot)}", size=9, bold=True, color="1D343D", after=4)

        block(urg, 1, "Urgent / en retard")
        block(mo, 2, f"À contrôler – bon et ordre de paiement pour le MO – pour ordre de paiement au vendredi {vend.strftime('%d.%m.%Y')}", afaire=True)
        block(eg, 3, "À contrôler – bon de paiement interne EG & ET")
        section(doc, "4. À clarifier")
        if clar:
            table(doc, ["Affaire", "Fournisseur", "N° facture", "TTC CHF", "Remarque"],
                  [[(f.get("affaire") or "")[:34], f.get("fournisseur", ""), f.get("no_facture", ""), chf(f.get("montant_ttc")), (f.get("remarque") or "–")[:140]] for f in clar],
                  [3.6, 3.0, 2.2, 2.0, 6.2], size=8.5)
        else:
            empty(doc, "Aucune facture à clarifier.")
        section(doc, "5. Validées ces 7 derniers jours")
        if recent:
            table(doc, ["Affaire", "Fournisseur", "N° facture", "TTC CHF", "Statut"],
                  [[(f.get("affaire") or "")[:40], f.get("fournisseur", ""), f.get("no_facture", ""), chf(f.get("montant_ttc")), f.get("statut", "")] for f in recent],
                  [5.0, 3.8, 2.8, 2.2, 3.2], size=8.5)
        else:
            empty(doc)

    para(doc, "Document généré automatiquement – procédures ASQ_ACSA_026.09 / ASQ_ACSA_026.10. Préparé par Clémence, assistante IA d'A. Clément SA : les montants et dates proviennent de la lecture des factures ; en cas de doute, la facture fait foi.", size=8, color="7F7F7F", before=8)

    kind = " direction" if data["type"] == "direction" else ""
    local, dom = data["email"].lower().split("@")
    assert dom == "a-clement.ch", "destinataire interne uniquement"
    who = local.replace(".", " ")
    base = f"ACFIN_Rapport_{who}_Synthèse factures{kind}" + (" TEST" if data.get("test") else "") + f"_{today.strftime('%Y.%m.%d')}"
    docx_path = os.path.join(out_dir, base + ".docx")
    doc.save(docx_path)
    word_swap_logo(docx_path, os.path.join(tpl_dir, "MExxx.xx_Base Word_qui_20xx.xx.xx.docx"))
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(out_dir, base + ".pdf")


if __name__ == "__main__":
    data = json.load(open(sys.argv[1], encoding="utf-8"))
    print(build(data, sys.argv[2], sys.argv[3]))
