# -*- coding: utf-8 -*-
"""
Tests de non-régression de l'extraction de données de factures.

Chaque cas reproduit la STRUCTURE d'une vraie facture rencontrée en
production (avec des données fictives) : lignes d'acompte négatives,
sous-totaux, colonnes de totaux séparées de leurs libellés par l'OCR,
numéros de TVA intracommunautaire, tickets de caisse...

Lancement :  python -m pytest tests/  (ou simplement : python tests/test_extraction.py)
"""

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import traitement_factures as moteur  # noqa: E402
from traitement_factures import (  # noqa: E402
    convertir_montant,
    extraire_date_facture,
    extraire_montants,
    identifier_client,
    identifier_fournisseur,
    nettoyer_nom_fichier,
)


# ---------------------------------------------------------------------------
# Conversion des montants
# ---------------------------------------------------------------------------

def test_convertir_montant():
    assert convertir_montant("1 234,56") == 1234.56
    assert convertir_montant("1.234,56") == 1234.56
    assert convertir_montant("1,234.56") == 1234.56
    assert convertir_montant("1234.56 €") == 1234.56
    assert convertir_montant("4 526") == 4526.0
    assert convertir_montant("") is None
    assert convertir_montant("abc") is None


# ---------------------------------------------------------------------------
# Facture native "propre" : totaux libellés, plusieurs taux de TVA
# ---------------------------------------------------------------------------

FACTURE_STANDARD = """FOURNITOUT SARL
Facture N° 2026-0654
Date de facture : 28/06/2026
Client : SARL Dupont Consulting
Total HT : 1 200,00 €
TVA 20% : 220,00 €
TVA 5,5 % : 11,00 €
Total TTC : 1 431,00 €
Net à payer : 1 431,00 €
"""


def test_facture_standard():
    m = extraire_montants(FACTURE_STANDARD)
    assert m["total_ht"] == 1200.00
    assert m["tva"] == {"20": 220.00, "10": None, "5.5": 11.00}
    assert m["total_ttc"] == 1431.00
    assert extraire_date_facture(FACTURE_STANDARD) == "2026-06-28"
    clients = [("SARL Dupont Consulting", "SARL Dupont Consulting")]
    assert identifier_client(FACTURE_STANDARD, clients) == "SARL Dupont Consulting"
    assert identifier_fournisseur(FACTURE_STANDARD) == "FOURNITOUT SARL"


# ---------------------------------------------------------------------------
# Facture d'acompte : "Montant HT - 470 180,00 €" dans le tableau d'articles
# ne doit PAS être pris pour le total ; seul le bloc du bas fait foi.
# ---------------------------------------------------------------------------

FACTURE_ACOMPTE = """BATIPRO 40
En date du : 08/06/2026
1.2 Montant HT - 470 180,00 € 1 u 0,00 € 0,00 €
1.3 Acompte HT N°1 devis 1 u 0,00 € 0,00 €
- 152 154,00 €
Total HT 60000,00 €
TVA 20 % 12000,00 €
Total TTC 72000,00 €
Net à payer 72000,00 €
TVA N° FR13882796386
"""


def test_facture_acompte_montants_negatifs():
    m = extraire_montants(FACTURE_ACOMPTE)
    assert m["total_ht"] == 60000.00
    assert m["tva"]["20"] == 12000.00
    assert m["total_ttc"] == 72000.00


# ---------------------------------------------------------------------------
# Sous-total + remise : le "Sous-total HT" ne doit pas masquer le "Total HT".
# ---------------------------------------------------------------------------

FACTURE_REMISE = """CLIMAPLUS
Sous-total HT 25 000,00 €
Remise HT -5 000,00 €
Total HT 20 000,00 €
TVA 20 % 4 000,00 €
Total TTC 24 000,00 €
"""


def test_facture_sous_total():
    m = extraire_montants(FACTURE_REMISE)
    assert m["total_ht"] == 20000.00
    assert m["total_ttc"] == 24000.00


# ---------------------------------------------------------------------------
# OCR : colonne de totaux SANS libellés (labels et montants séparés).
# Le plus grand montant "seul sur sa ligne" est le TTC.
# Le numéro SIREN (sans décimales) ne doit pas être pris pour un montant.
# ---------------------------------------------------------------------------

FACTURE_OCR_COLONNES = """ELECPLUS SASU
TOTAL HT :
TOTAL TVA:
TOTAL TTC:
995236478
1150,00 €
45,00 €
1195,00 €
0,00 €
1195,00 €
"""


def test_ocr_colonne_sans_libelles():
    m = extraire_montants(FACTURE_OCR_COLONNES)
    assert m["total_ttc"] == 1195.00


# ---------------------------------------------------------------------------
# Le numéro de TVA intracommunautaire ne doit jamais être lu comme un montant.
# ---------------------------------------------------------------------------

FACTURE_TVA_INTRACOM = """GARAGE MARTIN
TVA Intra-com : FR66405378829
Client: 03092 TTC : 1 967,18
"""


def test_ttc_nu_et_tva_intracom():
    m = extraire_montants(FACTURE_TVA_INTRACOM)
    assert m["total_ttc"] == 1967.18
    assert m["tva"] == {"20": None, "10": None, "5.5": None}


# ---------------------------------------------------------------------------
# Avis des finances publiques : "SOMME À PAYER".
# ---------------------------------------------------------------------------

def test_somme_a_payer():
    texte = "SOMME À PAYER : 100,32 Euro(s) TALON DE PAIEMENT"
    assert extraire_montants(texte)["total_ttc"] == 100.32


# ---------------------------------------------------------------------------
# Ticket de caisse : paiement par carte bancaire, TVA sans taux sur la ligne.
# ---------------------------------------------------------------------------

FACTURE_TICKET = """BRICOMARCHE
Carte-bancaire 122.38 EUR
Total TVA: 20.39 HT : 101.93
"""


def test_ticket_carte_bancaire():
    m = extraire_montants(FACTURE_TICKET)
    assert m["total_ttc"] == 122.38
    assert m["tva"]["20"] == 20.39  # taux déduit du ratio TVA/HT absent -> 20 %


# ---------------------------------------------------------------------------
# TVA sans taux affiché : déduction du taux par le ratio TVA / HT.
# ---------------------------------------------------------------------------

def test_tva_sans_taux_ratio_10():
    texte = "Montant HT 500.00\nTotal TVA 50.00\nTotal TTC 550.00"
    m = extraire_montants(texte)
    assert m["tva"] == {"20": None, "10": 50.0, "5.5": None}


# ---------------------------------------------------------------------------
# "TotalHT" collé (OCR) et date en toutes lettres.
# ---------------------------------------------------------------------------

def test_ocr_total_colle():
    texte = "TotalHT 600,00\nTVA 120,00\nTotal TTC 720,00"
    m = extraire_montants(texte)
    assert m["total_ht"] == 600.00
    assert m["total_ttc"] == 720.00


def test_date_en_lettres():
    assert extraire_date_facture("Facture émise le 12 février 2026") == "2026-02-12"


def test_date_iso():
    assert extraire_date_facture("Date: 2026-07-03") == "2026-07-03"


# ---------------------------------------------------------------------------
# Noms de fichiers/dossiers
# ---------------------------------------------------------------------------

def test_identifier_client_variantes():
    # Toutes les variantes doivent renvoyer le même nom canonique de dossier.
    clients = [("ICG 40", "ICG 40"), ("ICG40", "ICG 40"), ("SARL ICG 40", "ICG 40")]
    assert identifier_client("Facturé à : ICG40, Gabarret", clients) == "ICG 40"
    assert identifier_client("SARL ICG 40 — chantier Tesla", clients) == "ICG 40"
    assert identifier_client("Client inconnu SAS", clients) is None


def test_nettoyer_nom_fichier():
    assert nettoyer_nom_fichier("SARL Dupont & Cie / Été") == "SARL_Dupont_Cie_Ete"
    assert nettoyer_nom_fichier("") == "sans_nom"


# ---------------------------------------------------------------------------
# Tickets carburant (OCR) : TOT TIC, % lu comme &, Net = HT, MONTANT REEL
# ---------------------------------------------------------------------------

TICKET_CARBURANT = """STATION INTERMARCHE
Date 17-06-2026 14:19:45
Pompe 4 Gasoi 1
TOT TIC € 19.99
TVA. 20.00 & € 3,33
Net € 16.66
MONTANT REEL
EUR 19.99
"""


def test_ticket_carburant_ocr():
    m = extraire_montants(TICKET_CARBURANT)
    assert m["total_ttc"] == 19.99          # "TOT TIC" = TOT TTC déformé
    assert m["tva"]["20"] == 3.33           # "%" lu comme "&"
    assert m["total_ht"] == 16.66           # "Net" nu = HT sur les tickets
    assert extraire_date_facture(TICKET_CARBURANT) == "2026-06-17"


def test_montant_reel_ligne_suivante():
    texte = "CARREFOUR\nMONTANT REEL\nEUR 104.27\nDEBIT"
    assert extraire_montants(texte)["total_ttc"] == 104.27


def test_taux_jamais_pris_pour_montant():
    # "TVA. 20.00 &" : 20.00 est un TAUX, pas un montant de TVA.
    m = extraire_montants("TVA. 20.00 &\nrien d'autre")
    assert m["tva"] == {"20": None, "10": None, "5.5": None}


# ---------------------------------------------------------------------------
# "TOTAL" nu + chiffre parasite (€ lu comme 6), et avis FPS "est égal :"
# ---------------------------------------------------------------------------

def test_total_nu_et_chiffre_parasite():
    texte = ("FACTURE\nDÉSIGNATION MONTANT\nChantier Bazin 450,00\n"
             "TOTAL 2 250,00 6\nTVA non applicable art 293B")
    m = extraire_montants(texte)
    assert m["total_ttc"] == 2250.00        # le "6" parasite est ignoré


def test_avis_fps_est_egal():
    texte = "Le montante au IFRS du'est égal :45 eu"
    assert extraire_montants(texte)["total_ttc"] == 45.0


def test_date_annee_invraisemblable_rejetee():
    texte = "DATE D'ÉMISSION 17 juin 2028\nDATE DE LIVRAISON 17 juin 2026"
    assert extraire_date_facture(texte) == "2026-06-17"


# ---------------------------------------------------------------------------
# Doublons : fichiers identiques supprimés, tableau Excel nettoyé
# ---------------------------------------------------------------------------

def test_suppression_doublons():
    import pandas as pd

    with tempfile.TemporaryDirectory() as dossier:
        base = Path(dossier) / "Factures_Clients"
        excel = Path(dossier) / "tableau.xlsx"
        (base / "ICG_40").mkdir(parents=True)
        (base / "AUTRE").mkdir()

        # Deux fichiers au contenu identique (doublon inter-dossiers)
        # et un fichier distinct.
        (base / "ICG_40" / "originale.pdf").write_bytes(b"CONTENU-A")
        (base / "AUTRE" / "copie.pdf").write_bytes(b"CONTENU-A")
        (base / "ICG_40" / "unique.pdf").write_bytes(b"CONTENU-B")

        pd.DataFrame([
            {"Date de traitement": "t1", "Date de facture": "2026-06-01",
             "Client": "ICG 40", "Fournisseur": "X",
             "Nom du fichier": "originale.pdf", "Total HT": 100,
             "TVA 20%": 20, "TVA 10%": None, "TVA 5.5%": None,
             "Total TTC": 120},
            {"Date de traitement": "t2", "Date de facture": "2026-06-01",
             "Client": "AUTRE", "Fournisseur": "X",
             "Nom du fichier": "copie.pdf", "Total HT": 100,
             "TVA 20%": 20, "TVA 10%": None, "TVA 5.5%": None,
             "Total TTC": 120},
            {"Date de traitement": "t3", "Date de facture": "2026-06-02",
             "Client": "ICG 40", "Fournisseur": "Y",
             "Nom du fichier": "unique.pdf", "Total HT": 50,
             "TVA 20%": 10, "TVA 10%": None, "TVA 5.5%": None,
             "Total TTC": 60},
        ]).to_excel(excel, index=False)

        config = {"dossier_base": base, "fichier_excel": excel}
        fichiers, lignes = moteur.supprimer_doublons(config)

        assert fichiers == 1                       # une des deux copies part
        restants = {p.name for p in base.rglob("*.pdf")}
        # Il reste exactement UN exemplaire du contenu dupliqué + l'unique.
        assert "unique.pdf" in restants and len(restants) == 2
        supprime = ({"originale.pdf", "copie.pdf"} - restants).pop()
        df = pd.read_excel(excel)
        # La ligne Excel du fichier supprimé est retirée, l'autre subsiste.
        assert len(df) == 2 and supprime not in set(df["Nom du fichier"])
        # Le registre d'empreintes est reconstruit avec les fichiers restants.
        assert len(moteur.charger_empreintes(base)) == 2


def test_reclassement_facture():
    import pandas as pd

    with tempfile.TemporaryDirectory() as dossier:
        base = Path(dossier) / "Factures_Clients"
        excel = Path(dossier) / "tableau.xlsx"
        (base / "_A_CLASSER").mkdir(parents=True)
        pdf = base / "_A_CLASSER" / "2026-06-01_X_120.00.pdf"
        pdf.write_bytes(b"CONTENU-C")
        moteur.sauvegarder_empreintes(
            base, {moteur.empreinte_pdf(pdf): "_A_CLASSER/" + pdf.name})
        pd.DataFrame([
            {"Date de traitement": "t1", "Date de facture": "2026-06-01",
             "Client": "_A_CLASSER", "Fournisseur": "X",
             "Nom du fichier": pdf.name, "Total HT": 100, "TVA 20%": 20,
             "TVA 10%": None, "TVA 5.5%": None, "Total TTC": 120},
        ]).to_excel(excel, index=False)

        config = {"dossier_base": base, "fichier_excel": excel}
        destination = moteur.reclasser_facture(config, pdf, "ICG 40")

        assert destination.parent.name == "ICG_40" and destination.exists()
        assert not pdf.exists()
        df = pd.read_excel(excel)
        assert df.iloc[0]["Client"] == "ICG 40"
        empreintes = moteur.charger_empreintes(base)
        assert list(empreintes.values()) == [f"ICG_40/{destination.name}"]


# ---------------------------------------------------------------------------
# Export vers le classeur TVA officiel (modele_tva.xlsx)
# ---------------------------------------------------------------------------

def test_export_classeur_tva():
    import openpyxl
    import pandas as pd

    import export_tva

    racine = Path(__file__).resolve().parent.parent
    with tempfile.TemporaryDirectory() as dossier:
        excel = Path(dossier) / "tableau.xlsx"
        pd.DataFrame([
            # Vente : ICG 40 émet la facture (fournisseur = client).
            {"Date de traitement": "t", "Date de facture": "2026-06-08",
             "Client": "ICG 40", "Fournisseur": "ICG_40_FACTURE",
             "Nom du fichier": "vente1.pdf", "Total HT": 60000.0,
             "TVA 20%": 12000.0, "TVA 10%": None, "TVA 5.5%": None,
             "Total TTC": 72000.0},
            # Achat fournisseur externe à 20 %.
            {"Date de traitement": "t", "Date de facture": "2026-06-01",
             "Client": "ICG 40", "Fournisseur": "POINT_P",
             "Nom du fichier": "achat1.pdf", "Total HT": 1384.83,
             "TVA 20%": 276.97, "TVA 10%": None, "TVA 5.5%": None,
             "Total TTC": 1661.80},
            # Achat sans TVA détectée -> code 0.
            {"Date de traitement": "t", "Date de facture": "2026-06-12",
             "Client": "ICG 40", "Fournisseur": "FINANCES_PUBLIQUES",
             "Nom du fichier": "achat2.pdf", "Total HT": None,
             "TVA 20%": None, "TVA 10%": None, "TVA 5.5%": None,
             "Total TTC": 100.32},
            # Autre mois : ne doit PAS apparaître dans l'export de juin.
            {"Date de traitement": "t", "Date de facture": "2026-05-20",
             "Client": "ICG 40", "Fournisseur": "EURO_PNEU",
             "Nom du fichier": "mai.pdf", "Total HT": 1639.32,
             "TVA 20%": 327.86, "TVA 10%": None, "TVA 5.5%": None,
             "Total TTC": 1967.18},
        ]).to_excel(excel, index=False)

        config = {"fichier_excel": excel,
                  "modele_tva": racine / "modele_tva.xlsx",
                  "dossier_exports": Path(dossier) / "Exports_TVA"}

        assert export_tva.mois_disponibles(excel) == ["2026-06", "2026-05"]
        destination, exportees, ignorees = export_tva.exporter_mois(
            config, "2026-06")
        assert exportees == 3 and ignorees == 0

        classeur = openpyxl.load_workbook(destination)
        achats = classeur["TVA Deduct"]
        ventes = classeur["TVA Collectée"]
        # Achats triés par date : POINT_P (code 2) puis FINANCES (code 0).
        assert achats["D8"].value == "POINT_P"
        assert achats["F8"].value == 2 and achats["J8"].value == 1661.80
        assert achats["D9"].value == "FINANCES_PUBLIQUES"
        assert achats["F9"].value == 0 and achats["J9"].value == 100.32
        # La formule du modèle est préservée (HT calculé par Excel).
        assert str(achats["E8"].value).startswith("=IF(")
        # Vente en TVA Collectée avec le bon code et le bon TTC.
        assert ventes["A7"].value == "ICG 40"
        assert ventes["D7"].value == 2 and ventes["E7"].value == 72000.0
        # Titres mensualisés.
        assert "06/2026" in classeur["Résumé"]["B2"].value


if __name__ == "__main__":
    # Exécution sans pytest : lance toutes les fonctions test_*.
    erreurs = 0
    for nom, fonction in sorted(globals().items()):
        if nom.startswith("test_") and callable(fonction):
            try:
                fonction()
                print(f"  OK  {nom}")
            except AssertionError as e:
                erreurs += 1
                print(f"  KO  {nom} : {e}")
    print(f"\n{'Tous les tests passent.' if not erreurs else f'{erreurs} test(s) en échec.'}")
    sys.exit(1 if erreurs else 0)
