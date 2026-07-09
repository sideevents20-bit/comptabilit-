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
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

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
