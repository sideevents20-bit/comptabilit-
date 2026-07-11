# -*- coding: utf-8 -*-
"""
export_tva.py
=============

Export des factures traitées vers le classeur TVA mensuel officiel
(modele_tva.xlsx), dont la structure est imposée :

    - feuille « TVA Deduct »   : factures d'ACHAT (fournisseurs),
      lignes 8 à 21 (le sous-total ligne 22, les dépenses récurrentes
      lignes 24-30 et le Total General ligne 31 sont préservés) ;
    - feuille « TVA Collectée »: factures de VENTE, lignes 7 à 37
      (Total ligne 38) ;
    - feuille « Résumé »       : formules alimentées par les deux autres.

Le modèle calcule lui-même HT et TVA à partir du code TVA et du montant
TTC ; on n'écrit donc que les colonnes de saisie :
    code 2 = 20 %   |   code 1 = 10 %   |   code 55 = 5,5 %   |   0 = exonéré
"""

import logging
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

from traitement_factures import normaliser

logger = logging.getLogger("factures")

# Emplacements de saisie du modèle (bornes incluses).
PREMIERE_LIGNE_ACHATS, DERNIERE_LIGNE_ACHATS = 8, 21     # TVA Deduct
PREMIERE_LIGNE_VENTES, DERNIERE_LIGNE_VENTES = 7, 37     # TVA Collectée


def code_tva(ligne: pd.Series) -> int:
    """
    Code TVA du modèle pour une facture : celui du taux dont le montant
    de TVA est le plus élevé (2=20 %, 1=10 %, 55=5,5 %, 0=exonéré).
    Une facture multi-taux est signalée dans le journal.
    """
    montants = {
        2: float(ligne.get("TVA 20%") or 0),
        1: float(ligne.get("TVA 10%") or 0),
        55: float(ligne.get("TVA 5.5%") or 0),
    }
    presents = [code for code, montant in montants.items() if montant > 0]
    if not presents:
        return 0
    if len(presents) > 1:
        logger.warning(
            "%s : plusieurs taux de TVA sur la même facture — le modèle "
            "n'accepte qu'un code par ligne, le taux principal est retenu.",
            ligne.get("Nom du fichier"),
        )
    return max(presents, key=lambda code: montants[code])


def est_vente(ligne: pd.Series) -> bool:
    """
    Vente ou achat ? C'est une VENTE si l'émetteur de la facture (colonne
    Fournisseur) est le client lui-même (ex : ICG 40 émet la facture) ;
    sinon c'est un achat auprès d'un fournisseur externe.
    """
    client = normaliser(str(ligne.get("Client", ""))).replace(" ", "")
    fournisseur = normaliser(str(ligne.get("Fournisseur", ""))).replace(" ", "")
    fournisseur = fournisseur.replace("_", "")
    return bool(client) and client in fournisseur


def _date_cellule(valeur) -> datetime | str:
    """Convertit la date ISO du tableau global en datetime pour Excel."""
    try:
        return datetime.strptime(str(valeur)[:10], "%Y-%m-%d")
    except ValueError:
        return str(valeur)


def mois_disponibles(fichier_excel: Path) -> list[str]:
    """Mois (AAAA-MM) présents dans le tableau global, du plus récent."""
    if not fichier_excel.exists():
        return []
    df = pd.read_excel(fichier_excel)
    mois = (pd.to_datetime(df.get("Date de facture"), errors="coerce")
            .dt.strftime("%Y-%m").dropna().unique())
    return sorted(mois, reverse=True)


def exporter_mois(config: dict, mois: str) -> tuple[Path, int, int]:
    """
    Génère le classeur TVA du mois demandé (format « AAAA-MM ») à partir
    du tableau global et du modèle. Retourne (chemin_du_fichier,
    nb_factures_exportées, nb_factures_ignorées_faute_de_place).

    Le fichier Exports_TVA/TVA_AAAA-MM.xlsx est (ré)écrit à chaque appel :
    il reflète toujours l'état courant du tableau global.
    """
    modele = config["modele_tva"]
    if not modele.exists():
        raise FileNotFoundError(
            f"Modèle introuvable : {modele}. Placez modele_tva.xlsx à côté "
            "de l'application.")

    df = pd.read_excel(config["fichier_excel"])
    df["_mois"] = (pd.to_datetime(df["Date de facture"], errors="coerce")
                   .dt.strftime("%Y-%m"))
    selection = df[df["_mois"] == mois].copy()
    if selection.empty:
        raise ValueError(f"Aucune facture datée de {mois} dans le tableau.")
    selection = selection.sort_values("Date de facture")

    annee, numero_mois = mois.split("-")
    classeur = openpyxl.load_workbook(modele)

    # --- Titres des feuilles -------------------------------------------------
    classeur["TVA Deduct"]["A2"] = f"Liste Factures Achat - {numero_mois}/{annee}"
    classeur["TVA Collectée"]["A2"] = f"TVA COLLECTEE {numero_mois}/{annee}"
    classeur["Résumé"]["B2"] = f"Résumé Mouvements TVA {numero_mois}/{annee}"

    achats = classeur["TVA Deduct"]
    ventes = classeur["TVA Collectée"]
    ligne_achat, ligne_vente = PREMIERE_LIGNE_ACHATS, PREMIERE_LIGNE_VENTES
    numero_registre, exportees, ignorees = 1, 0, 0

    for _, facture in selection.iterrows():
        ttc = float(facture.get("Total TTC") or 0)
        date = _date_cellule(facture.get("Date de facture"))
        code = code_tva(facture)

        if est_vente(facture):
            if ligne_vente > DERNIERE_LIGNE_VENTES:
                ignorees += 1
                continue
            ventes[f"A{ligne_vente}"] = facture.get("Client")
            cellule_date = ventes[f"C{ligne_vente}"]
            cellule_date.value = date
            cellule_date.number_format = "dd/mm/yy"
            ventes[f"D{ligne_vente}"] = code
            ventes[f"E{ligne_vente}"] = round(ttc, 2)
            ventes[f"O{ligne_vente}"] = str(facture.get("Nom du fichier", ""))
            ligne_vente += 1
        else:
            if ligne_achat > DERNIERE_LIGNE_ACHATS:
                ignorees += 1
                continue
            achats[f"A{ligne_achat}"] = numero_registre
            cellule_date = achats[f"C{ligne_achat}"]
            cellule_date.value = date
            cellule_date.number_format = "dd/mm/yy"
            achats[f"D{ligne_achat}"] = facture.get("Fournisseur")
            achats[f"F{ligne_achat}"] = code
            achats[f"J{ligne_achat}"] = round(ttc, 2)
            achats[f"K{ligne_achat}"] = str(facture.get("Nom du fichier", ""))
            numero_registre += 1
            ligne_achat += 1
        exportees += 1

    dossier = config["dossier_exports"]
    dossier.mkdir(parents=True, exist_ok=True)
    destination = dossier / f"TVA_{mois}.xlsx"
    classeur.save(destination)

    if ignorees:
        logger.warning(
            "Export %s : %d facture(s) non reportée(s) faute d'emplacements "
            "libres dans le modèle (14 achats / 31 ventes max par mois) — "
            "à saisir manuellement.", mois, ignorees)
    logger.info("Classeur TVA généré : %s (%d facture(s)).",
                destination, exportees)
    return destination, exportees, ignorees
