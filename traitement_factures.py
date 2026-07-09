#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
traitement_factures.py
======================

Script d'automatisation du traitement des factures reçues par email.

Pipeline complet :
    1. Connexion IMAP et relève des emails non lus (mots-clés dans l'objet
       ou pièces jointes PDF).
    2. Téléchargement des pièces jointes PDF dans un dossier temporaire.
    3. Extraction des données de la facture (client, fournisseur, HT,
       TVA par taux, TTC, date) via pdfplumber + expressions régulières.
    4. Classement du PDF dans le dossier du client, renommé au format
       YYYY-MM-DD_Fournisseur_MontantTTC.pdf.
    5. Ajout d'une ligne dans le tableau Excel global de TVA (pandas).

Configuration : voir le fichier .env (identifiants IMAP, chemins, etc.)
et le fichier clients.txt (liste des clients du centre d'affaires).

Usage :
    python traitement_factures.py
"""

import email
import imaplib
import logging
import os
import re
import shutil
import sys
import unicodedata
from datetime import datetime
from email.header import decode_header
from email.message import Message
from email.utils import parseaddr
from pathlib import Path

import pandas as pd
import pdfplumber
from dotenv import load_dotenv

# --- OCR optionnel (pour les factures scannées) ----------------------------
# Nécessite : pip install pytesseract pypdfium2 Pillow
# + le logiciel Tesseract avec le pack français (voir README).
try:
    import pypdfium2 as pdfium
    import pytesseract

    OCR_DISPONIBLE = True
except ImportError:
    OCR_DISPONIBLE = False

# ---------------------------------------------------------------------------
# 0. CONFIGURATION & LOGGING
# ---------------------------------------------------------------------------

logger = logging.getLogger("factures")

# Dossier utilisé quand aucun client connu n'est identifié dans la facture.
DOSSIER_NON_CLASSE = "_A_CLASSER"

# Colonnes du tableau Excel global (l'ordre est conservé à l'écriture).
COLONNES_EXCEL = [
    "Date de traitement",
    "Date de facture",
    "Client",
    "Fournisseur",
    "Nom du fichier",
    "Total HT",
    "TVA 20%",
    "TVA 10%",
    "TVA 5.5%",
    "Total TTC",
]

# Mois français -> numéro, pour les dates écrites en toutes lettres.
MOIS_FR = {
    "janvier": 1, "fevrier": 2, "février": 2, "mars": 3, "avril": 4,
    "mai": 5, "juin": 6, "juillet": 7, "aout": 8, "août": 8,
    "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12,
    "décembre": 12,
}


def configurer_logging() -> None:
    """Configure des logs lisibles dans le terminal + fichier de suivi."""
    fmt = "%(asctime)s | %(levelname)-8s | %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler("traitement_factures.log", encoding="utf-8"),
        ],
    )


def charger_configuration() -> dict:
    """
    Charge la configuration depuis le fichier .env.

    Lève une erreur explicite si un identifiant obligatoire est absent,
    afin de ne jamais tenter une connexion avec des valeurs vides.
    """
    load_dotenv()

    obligatoires = ["IMAP_HOST", "IMAP_USER", "IMAP_PASSWORD"]
    manquants = [cle for cle in obligatoires if not os.getenv(cle)]
    if manquants:
        raise RuntimeError(
            f"Variables manquantes dans le fichier .env : {', '.join(manquants)}. "
            "Copiez .env.example vers .env et renseignez vos identifiants."
        )

    config = {
        "imap_host": os.getenv("IMAP_HOST"),
        "imap_port": int(os.getenv("IMAP_PORT", "993")),
        "imap_user": os.getenv("IMAP_USER"),
        "imap_password": os.getenv("IMAP_PASSWORD"),
        "imap_folder": os.getenv("IMAP_FOLDER", "INBOX"),
        "dossier_base": Path(os.getenv("DOSSIER_BASE", "./Factures_Clients")),
        "dossier_temp": Path(os.getenv("DOSSIER_TEMP", "./_temp_factures")),
        "fichier_excel": Path(os.getenv("FICHIER_EXCEL", "./tableau_tva_global.xlsx")),
        "fichier_clients": Path(os.getenv("FICHIER_CLIENTS", "./clients.txt")),
        "mots_cles": [
            m.strip().lower()
            for m in os.getenv("MOTS_CLES_OBJET", "facture,invoice").split(",")
            if m.strip()
        ],
        # Adresse dédiée aux factures (ex : alias Gmail "+factures").
        # Si renseignée, seuls les emails envoyés À cette adresse sont relevés.
        "filtre_destinataire": os.getenv("FILTRE_DESTINATAIRE", "").strip(),
    }
    return config


def charger_clients(fichier_clients: Path) -> list[tuple[str, str]]:
    """
    Charge la liste des clients ('#' pour commenter). Deux syntaxes :

        ICG 40                          -> nom unique
        ICG 40 = ICG40, SARL ICG 40     -> nom du dossier + variantes

    Retourne une liste de couples (variante_cherchée, nom_canonique) :
    chaque variante trouvée dans un PDF rattache la facture au même
    dossier client (le nom canonique, à gauche du '=').
    """
    if not fichier_clients.exists():
        logger.warning(
            "Fichier clients introuvable (%s) : toutes les factures iront "
            "dans le dossier '%s'.", fichier_clients, DOSSIER_NON_CLASSE
        )
        return []

    clients = []
    for ligne in fichier_clients.read_text(encoding="utf-8").splitlines():
        ligne = ligne.strip()
        if not ligne or ligne.startswith("#"):
            continue
        if "=" in ligne:
            canonique, variantes = ligne.split("=", 1)
            canonique = canonique.strip()
            clients.append((canonique, canonique))
            for variante in variantes.split(","):
                if variante.strip():
                    clients.append((variante.strip(), canonique))
        else:
            clients.append((ligne, ligne))
    nb_canoniques = len({c for _, c in clients})
    logger.info(
        "%d client(s) chargé(s) (%d variante(s)) depuis %s",
        nb_canoniques, len(clients), fichier_clients,
    )
    return clients


# ---------------------------------------------------------------------------
# 1. CONNEXION IMAP & RELÈVE DES EMAILS
# ---------------------------------------------------------------------------

def connecter_imap(config: dict) -> imaplib.IMAP4_SSL:
    """Ouvre une connexion IMAP sécurisée (SSL) et sélectionne la boîte."""
    logger.info("Connexion IMAP à %s:%s ...", config["imap_host"], config["imap_port"])
    connexion = imaplib.IMAP4_SSL(config["imap_host"], config["imap_port"])
    connexion.login(config["imap_user"], config["imap_password"])
    statut, _ = connexion.select(config["imap_folder"])
    if statut != "OK":
        raise RuntimeError(f"Impossible de sélectionner la boîte {config['imap_folder']}")
    logger.info("Connecté à la boîte '%s'.", config["imap_folder"])
    return connexion


def decoder_entete(valeur: str | None) -> str:
    """Décode un en-tête email (objet, nom de fichier) potentiellement encodé."""
    if not valeur:
        return ""
    morceaux = decode_header(valeur)
    resultat = ""
    for texte, charset in morceaux:
        if isinstance(texte, bytes):
            resultat += texte.decode(charset or "utf-8", errors="replace")
        else:
            resultat += texte
    return resultat


def contient_pdf(message: Message) -> bool:
    """Indique si l'email contient au moins une pièce jointe PDF."""
    for partie in message.walk():
        nom = decoder_entete(partie.get_filename())
        if partie.get_content_type() == "application/pdf" or nom.lower().endswith(".pdf"):
            return True
    return False


def telecharger_pieces_jointes_pdf(message: Message, dossier_temp: Path) -> list[Path]:
    """
    Télécharge toutes les pièces jointes PDF d'un email dans le dossier
    temporaire et retourne la liste des chemins créés.
    """
    chemins = []
    for partie in message.walk():
        nom_fichier = decoder_entete(partie.get_filename())
        est_pdf = (
            partie.get_content_type() == "application/pdf"
            or nom_fichier.lower().endswith(".pdf")
        )
        if not est_pdf or not nom_fichier:
            continue

        contenu = partie.get_payload(decode=True)
        if not contenu:
            continue

        # Nom de fichier sûr + anti-collision dans le dossier temporaire.
        nom_sur = nettoyer_nom_fichier(nom_fichier)
        chemin = dossier_temp / nom_sur
        compteur = 1
        while chemin.exists():
            chemin = dossier_temp / f"{Path(nom_sur).stem}_{compteur}.pdf"
            compteur += 1

        chemin.write_bytes(contenu)
        chemins.append(chemin)
        logger.info("Pièce jointe téléchargée : %s", chemin.name)
    return chemins


def relever_emails_factures(connexion: imaplib.IMAP4_SSL, config: dict) -> list[dict]:
    """
    Recherche les emails NON LUS pertinents et télécharge leurs PDF.

    Un email est pertinent si son objet contient un mot-clé (facture,
    invoice, ...) OU s'il contient une pièce jointe PDF.

    Retourne une liste de dictionnaires :
        {"id": ..., "objet": ..., "expediteur": ..., "pdfs": [Path, ...]}

    Si FILTRE_DESTINATAIRE est renseigné dans le .env (adresse email
    dédiée aux factures), la recherche est restreinte aux emails envoyés
    à cette adresse (en-tête To), et le filtre mot-clé/PDF est conservé
    comme garde-fou.

    NB : la lecture se fait avec BODY.PEEK pour ne PAS marquer l'email
    comme lu ; il n'est marqué lu qu'après traitement réussi.
    """
    criteres = ["UNSEEN"]
    if config["filtre_destinataire"]:
        criteres += ["TO", f'"{config["filtre_destinataire"]}"']
        logger.info(
            "Filtre actif : uniquement les emails adressés à %s.",
            config["filtre_destinataire"],
        )
    statut, donnees = connexion.search(None, *criteres)
    if statut != "OK":
        raise RuntimeError("La recherche IMAP des emails non lus a échoué.")

    ids = donnees[0].split()
    logger.info("%d email(s) non lu(s) trouvé(s).", len(ids))

    emails_pertinents = []
    for msg_id in ids:
        statut, contenu = connexion.fetch(msg_id, "(BODY.PEEK[])")
        if statut != "OK" or not contenu or contenu[0] is None:
            logger.warning("Impossible de lire l'email id=%s, ignoré.", msg_id)
            continue

        message = email.message_from_bytes(contenu[0][1])
        objet = decoder_entete(message.get("Subject"))
        nom_exp, adresse_exp = parseaddr(decoder_entete(message.get("From")))
        expediteur = nom_exp or adresse_exp

        objet_pertinent = any(mc in objet.lower() for mc in config["mots_cles"])
        pdf_present = contient_pdf(message)

        if not (objet_pertinent or pdf_present):
            logger.info("Email ignoré (ni mot-clé ni PDF) : « %s »", objet)
            continue

        pdfs = telecharger_pieces_jointes_pdf(message, config["dossier_temp"])
        if not pdfs:
            logger.warning("Email « %s » : aucun PDF téléchargeable, ignoré.", objet)
            continue

        emails_pertinents.append(
            {"id": msg_id, "objet": objet, "expediteur": expediteur, "pdfs": pdfs}
        )
        logger.info("Email retenu : « %s » (%d PDF).", objet, len(pdfs))

    return emails_pertinents


def marquer_comme_lu(connexion: imaplib.IMAP4_SSL, msg_id: bytes) -> None:
    """Marque un email comme lu (une fois toutes ses factures traitées)."""
    connexion.store(msg_id, "+FLAGS", "\\Seen")


# ---------------------------------------------------------------------------
# 2. ANALYSE DU PDF (EXTRACTION DE DONNÉES)
# ---------------------------------------------------------------------------

def extraire_texte_pdf(chemin_pdf: Path) -> str:
    """
    Extrait le texte du PDF avec pdfplumber ; si le PDF est un scan
    (pas de couche texte), bascule automatiquement sur l'OCR Tesseract.
    """
    texte = []
    with pdfplumber.open(chemin_pdf) as pdf:
        for page in pdf.pages:
            texte.append(page.extract_text() or "")
    resultat = "\n".join(texte)

    # Moins de 30 caractères utiles = très probablement un scan.
    if len(resultat.strip()) >= 30:
        return resultat

    if not OCR_DISPONIBLE:
        logger.warning(
            "%s : PDF scanné et OCR indisponible. Installez pytesseract, "
            "pypdfium2 et Tesseract (voir README).", chemin_pdf.name
        )
        return resultat

    logger.info("%s : PDF scanné détecté, lecture par OCR...", chemin_pdf.name)
    return ocr_pdf(chemin_pdf)


def ocr_pdf(chemin_pdf: Path, dpi: int = 300) -> str:
    """
    OCR d'un PDF scanné : chaque page est convertie en image (pypdfium2)
    puis lue par Tesseract en français.
    """
    texte = []
    document = pdfium.PdfDocument(str(chemin_pdf))
    try:
        for page in document:
            image = page.render(scale=dpi / 72).to_pil()
            texte.append(pytesseract.image_to_string(image, lang="fra"))
    finally:
        document.close()
    return "\n".join(texte)


def normaliser(texte: str) -> str:
    """Minuscules + suppression des accents, pour des comparaisons robustes."""
    texte = unicodedata.normalize("NFD", texte)
    texte = "".join(c for c in texte if unicodedata.category(c) != "Mn")
    return texte.lower()


def convertir_montant(brut: str) -> float | None:
    """
    Convertit un montant texte en float.

    Gère les formats français et anglo-saxons :
        "1 234,56"  -> 1234.56
        "1.234,56"  -> 1234.56
        "1,234.56"  -> 1234.56
        "1234.56 €" -> 1234.56
    """
    if not brut:
        return None
    # Suppression devise et espaces (y compris insécables).
    brut = re.sub(r"[€$\s  ]", "", brut).strip()
    if not brut:
        return None

    if "," in brut and "." in brut:
        # Le dernier séparateur rencontré est le séparateur décimal.
        if brut.rfind(",") > brut.rfind("."):
            brut = brut.replace(".", "").replace(",", ".")
        else:
            brut = brut.replace(",", "")
    elif "," in brut:
        # Virgule seule = décimale à la française.
        brut = brut.replace(",", ".")

    try:
        return round(float(brut), 2)
    except ValueError:
        return None


# Motif générique d'un montant : chiffres, espaces, points, virgules.
_MONTANT = r"([\d][\d\s  .,]*)"


def _dernier_montant(motif: str, texte: str) -> float | None:
    """
    Retourne le montant de la DERNIERE occurrence du motif dans le texte.

    Sur une facture, les totaux fiables sont dans le bloc recapitulatif en
    bas de page ; les memes libelles peuvent apparaitre plus haut dans le
    tableau des articles (ex : "Montant HT - 470 180,00" sur une ligne
    d'acompte, ou un "Sous-total HT" intermediaire) et donneraient un
    resultat faux si on prenait la premiere occurrence.
    """
    montant = None
    for m in re.finditer(motif, texte, re.IGNORECASE):
        valeur = convertir_montant(m.group(1))
        if valeur is not None:
            montant = valeur
    return montant


def _plus_grand_montant_isole(texte: str) -> float | None:
    """
    Dernier recours pour le Total TTC : montants seuls sur leur ligne.

    Sur beaucoup de factures scannées, l'OCR sépare la colonne des totaux
    de ses libellés : on obtient des lignes ne contenant qu'un montant
    ("489,79 EUR" / "97,96 EUR" / "587,75 EUR"). Le TTC étant le plus
    grand de ces totaux, on retourne le maximum. Seuls les montants
    formatés avec 2 décimales sont retenus, pour exclure les numéros
    SIREN/IBAN et autres nombres parasites.
    """
    motif = re.compile(
        r"^\s*(\d{1,3}(?:[\s\u202f\u00a0.]?\d{3})*[.,]\d{2})\s*(?:€|EUROS?|EUR)?\s*$",
        re.IGNORECASE,
    )
    montants = []
    for ligne in texte.splitlines():
        m = motif.match(ligne)
        if m:
            valeur = convertir_montant(m.group(1))
            if valeur:
                montants.append(valeur)
    return max(montants) if montants else None


def extraire_date_facture(texte: str) -> str | None:
    """
    Extrait la date de la facture et la retourne au format ISO YYYY-MM-DD.

    Cherche d'abord une date libellée ("Date de facture : ..."), puis à
    défaut la première date plausible du document. Formats gérés :
    JJ/MM/AAAA, JJ-MM-AAAA, JJ.MM.AAAA, AAAA-MM-JJ, "12 janvier 2025".
    """
    date_num = r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})"
    date_iso = r"(\d{4})-(\d{2})-(\d{2})"
    date_lettres = r"(\d{1,2})(?:er)?\s+([a-zéèûôî]+)\s+(\d{4})"

    # 1) Date explicitement libellée (prioritaire, évite les dates d'échéance).
    libelles = (
        r"(?:date\s+(?:de\s+(?:la\s+)?)?(?:facture|facturation|émission|emission)"
        r"|facture\s+du|émise?\s+le|emise?\s+le|le)\s*:?\s*"
    )
    for motif in (libelles + date_num, libelles + date_lettres):
        m = re.search(motif, texte, re.IGNORECASE)
        if m:
            date = _construire_date(m.groups())
            if date:
                return date

    # 2) Repli : première date trouvée dans le document.
    m = re.search(date_iso, texte)
    if m:
        annee, mois, jour = m.groups()
        return _construire_date((jour, mois, annee))
    for motif in (date_num, date_lettres):
        m = re.search(motif, texte, re.IGNORECASE)
        if m:
            date = _construire_date(m.groups())
            if date:
                return date
    return None


def _construire_date(groupes: tuple) -> str | None:
    """Valide un triplet (jour, mois, année) et le formate en YYYY-MM-DD."""
    jour_s, mois_s, annee_s = groupes
    try:
        mois = MOIS_FR.get(normaliser(mois_s)) if not mois_s.isdigit() else int(mois_s)
        if mois is None:
            return None
        jour, annee = int(jour_s), int(annee_s)
        if annee < 100:  # année sur 2 chiffres : "25" -> 2025
            annee += 2000
        return datetime(annee, mois, jour).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def identifier_client(texte: str, clients: list[tuple[str, str]]) -> str | None:
    """
    Identifie le client de la facture : première variante de clients.txt
    présente dans le texte du PDF (comparaison sans casse ni accents).
    Retourne le nom canonique du client (celui du dossier de classement).
    """
    texte_norm = normaliser(texte)
    for variante, canonique in clients:
        if normaliser(variante) in texte_norm:
            return canonique
    return None


def identifier_fournisseur(texte: str, repli: str = "") -> str:
    """
    Identifie le fournisseur (émetteur de la facture).

    Stratégie :
        1. Ligne libellée "Fournisseur :", "Émetteur :", "Vendeur :".
        2. Première ligne non vide du PDF (l'en-tête porte presque
           toujours le nom de l'émetteur).
        3. Repli : nom de l'expéditeur de l'email.
    """
    m = re.search(
        r"(?:fournisseur|émetteur|emetteur|vendeur|vendu\s+par)\s*:?\s*(.+)",
        texte, re.IGNORECASE,
    )
    if m:
        candidat = m.group(1).strip()
        if candidat:
            return candidat[:60]

    for ligne in texte.splitlines():
        ligne = ligne.strip()
        # On écarte les lignes qui sont manifestement un titre de document.
        if ligne and not re.match(r"^(facture|invoice|devis|n°|no\b)", ligne, re.IGNORECASE):
            return ligne[:60]

    return repli or "Fournisseur_inconnu"


def extraire_montants(texte: str) -> dict:
    """
    Extrait Total HT, TVA par taux (20 / 10 / 5.5 %) et Total TTC.

    Retourne : {"total_ht": float|None, "tva": {"20": ..., "10": ..., "5.5": ...},
                "total_ttc": float|None}
    """
    resultat = {"total_ht": None, "tva": {"20": None, "10": None, "5.5": None},
                "total_ttc": None}

    # --- Total HT -----------------------------------------------------------
    # Priorité au "Total HT" (bloc récapitulatif), en excluant les
    # "Sous-total HT" intermédiaires et les montants négatifs (acomptes,
    # remises) ; on prend toujours la DERNIÈRE occurrence, celle du bas
    # de la facture. "Montant HT" / "Total hors taxes" servent de repli.
    resultat["total_ht"] = _dernier_montant(
        r"(?<![a-zà-ü-])total\s*h\.?t\.?[^\d\n-]{0,20}" + _MONTANT, texte
    )
    if resultat["total_ht"] is None:
        resultat["total_ht"] = _dernier_montant(
            r"(?:montant\s*h\.?t\.?|total\s+hors\s+taxes?)[^\d\n-]{0,20}" + _MONTANT,
            texte,
        )

    # --- TVA par taux -------------------------------------------------------
    # Exemples couverts : "TVA 20% : 100,00", "TVA (5,5 %) 12.30",
    #                     "Montant TVA 10,00 % 45,00 €"
    # Dernière occurrence par taux : c'est celle du récapitulatif de TVA.
    for m in re.finditer(
        r"tva[^\d\n%]{0,15}\(?\s*(20|10|5[.,]5)(?:[.,]0{1,2})?\s*%\s*\)?"
        r"[^\d\n-]{0,20}" + _MONTANT,
        texte, re.IGNORECASE,
    ):
        taux = m.group(1).replace(",", ".")
        montant = convertir_montant(m.group(2))
        if montant is not None:
            resultat["tva"][taux] = montant

    # Cas d'une TVA sans taux affiché sur la même ligne : "Total TVA : 98,40".
    # Seuls espaces / ':' / '.' sont tolérés entre "TVA" et le montant, pour
    # ne jamais capturer un numéro de TVA intracommunautaire ("TVA : FR13...").
    if all(v is None for v in resultat["tva"].values()):
        montant_tva = _dernier_montant(
            r"(?:total\s+|montant\s+)?tva[\s:.]{1,10}" + _MONTANT, texte
        )
        if montant_tva is not None:
            # Sans taux explicite, on le déduit du ratio TVA / HT si possible,
            # sinon on suppose le taux normal de 20 %.
            taux_deduit = "20"
            if resultat["total_ht"] and montant_tva:
                ratio = round(montant_tva / resultat["total_ht"] * 100, 1)
                if abs(ratio - 10) <= 0.5:
                    taux_deduit = "10"
                elif abs(ratio - 5.5) <= 0.5:
                    taux_deduit = "5.5"
            resultat["tva"][taux_deduit] = montant_tva

    # --- Total TTC ----------------------------------------------------------
    # Cascade de replis, du libellé le plus fiable au moins fiable :
    #   1. "Total TTC" / "Montant TTC" (dernière occurrence)
    #   2. "Net à payer" / "Montant dû" / "Somme à payer"
    #   3. "TTC" seul (ex : ticket "Client: 03092 TTC : 1 967,18")
    #   4. Montants seuls sur leur ligne (colonne de totaux sans libellés,
    #      fréquent sur les PDF passés à l'OCR) : on prend le plus grand.
    resultat["total_ttc"] = _dernier_montant(
        r"(?:total\s*t\.?t\.?c\.?|montant\s*t\.?t\.?c\.?)[^\d\n-]{0,20}" + _MONTANT,
        texte,
    )
    if resultat["total_ttc"] is None:
        resultat["total_ttc"] = _dernier_montant(
            r"(?:net\s+[àa]\s+payer|total\s+[àa]\s+payer|montant\s+d[ûu]"
            r"|somme\s+[àa]\s+payer|carte\s*-?\s*bancaire)[^\d\n-]{0,20}" + _MONTANT,
            texte,
        )
    if resultat["total_ttc"] is None:
        # [ \t] et non \s : le montant doit être sur la MÊME ligne que "TTC"
        # (l'OCR sépare souvent libellés et montants en colonnes distinctes).
        resultat["total_ttc"] = _dernier_montant(
            r"\bt\.?t\.?c\.?\b[ \t:.]{1,10}" + _MONTANT, texte
        )
    if resultat["total_ttc"] is None:
        resultat["total_ttc"] = _plus_grand_montant_isole(texte)

    # --- Contrôle de cohérence HT + TVA = TTC (simple avertissement) --------
    ht, ttc = resultat["total_ht"], resultat["total_ttc"]
    total_tva = sum(v for v in resultat["tva"].values() if v is not None)
    if ht is not None and ttc is not None and total_tva:
        if abs((ht + total_tva) - ttc) > 0.05:
            logger.warning(
                "Incohérence détectée : HT (%.2f) + TVA (%.2f) != TTC (%.2f).",
                ht, total_tva, ttc,
            )

    return resultat


def analyser_facture(chemin_pdf: Path, clients: list[str], expediteur: str) -> dict:
    """
    Analyse complète d'un PDF de facture.

    Retourne un dictionnaire avec toutes les données extraites. Lève une
    ValueError si les données minimales (montant TTC) sont introuvables.
    """
    texte = extraire_texte_pdf(chemin_pdf)
    if not texte.strip():
        raise ValueError(
            "Aucun texte extrait (PDF scanné/image ? un OCR serait nécessaire)."
        )

    montants = extraire_montants(texte)
    if montants["total_ttc"] is None:
        # Sans TTC on tente HT + TVA pour ne pas bloquer le traitement.
        ht = montants["total_ht"]
        tva = sum(v for v in montants["tva"].values() if v is not None)
        if ht is not None:
            montants["total_ttc"] = round(ht + tva, 2)
        else:
            raise ValueError("Montant Total TTC introuvable dans le PDF.")

    date_facture = extraire_date_facture(texte)
    if not date_facture:
        logger.warning(
            "%s : date de facture introuvable, date du jour utilisée.",
            chemin_pdf.name,
        )
        date_facture = datetime.now().strftime("%Y-%m-%d")

    client = identifier_client(texte, clients)
    if not client:
        logger.warning(
            "%s : aucun client connu identifié, classement dans '%s'.",
            chemin_pdf.name, DOSSIER_NON_CLASSE,
        )

    return {
        "client": client or DOSSIER_NON_CLASSE,
        "fournisseur": identifier_fournisseur(texte, repli=expediteur),
        "date_facture": date_facture,
        "total_ht": montants["total_ht"],
        "tva_20": montants["tva"]["20"],
        "tva_10": montants["tva"]["10"],
        "tva_5_5": montants["tva"]["5.5"],
        "total_ttc": montants["total_ttc"],
    }


# ---------------------------------------------------------------------------
# 3. CLASSEMENT DANS LES DOSSIERS CLIENTS
# ---------------------------------------------------------------------------

def nettoyer_nom_fichier(nom: str) -> str:
    """Rend une chaîne utilisable comme nom de fichier/dossier."""
    nom = unicodedata.normalize("NFKD", nom)
    nom = "".join(c for c in nom if unicodedata.category(c) != "Mn")
    nom = re.sub(r"[^\w\s.\-]", "", nom)      # caractères spéciaux interdits
    nom = re.sub(r"\s+", "_", nom.strip())    # espaces -> underscores
    return nom or "sans_nom"


def classer_facture(chemin_pdf: Path, donnees: dict, dossier_base: Path) -> Path:
    """
    Déplace la facture du dossier temporaire vers le dossier du client,
    en la renommant : YYYY-MM-DD_Fournisseur_MontantTTC.pdf.

    Crée le dossier client s'il n'existe pas. Retourne le chemin final.
    """
    dossier_client = dossier_base / nettoyer_nom_fichier(donnees["client"])
    if not dossier_client.exists():
        dossier_client.mkdir(parents=True)
        logger.info("Dossier client créé : %s", dossier_client)

    nom_final = (
        f"{donnees['date_facture']}_"
        f"{nettoyer_nom_fichier(donnees['fournisseur'])}_"
        f"{donnees['total_ttc']:.2f}.pdf"
    )

    destination = dossier_client / nom_final
    compteur = 1
    while destination.exists():  # anti-écrasement en cas de doublon
        destination = dossier_client / f"{destination.stem.rsplit('_v', 1)[0]}_v{compteur}.pdf"
        compteur += 1

    shutil.move(str(chemin_pdf), str(destination))
    logger.info("Facture classée : %s", destination)
    return destination


# ---------------------------------------------------------------------------
# 4. GÉNÉRATION DU TABLEAU DE TVA (EXCEL)
# ---------------------------------------------------------------------------

def ajouter_ligne_excel(donnees: dict, nom_fichier: str, fichier_excel: Path) -> None:
    """
    Ajoute une ligne au tableau Excel global de TVA.

    Le fichier est créé avec ses en-têtes s'il n'existe pas encore ;
    sinon la ligne est ajoutée à la suite des données existantes.
    """
    nouvelle_ligne = {
        "Date de traitement": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Date de facture": donnees["date_facture"],
        "Client": donnees["client"],
        "Fournisseur": donnees["fournisseur"],
        "Nom du fichier": nom_fichier,
        "Total HT": donnees["total_ht"],
        "TVA 20%": donnees["tva_20"],
        "TVA 10%": donnees["tva_10"],
        "TVA 5.5%": donnees["tva_5_5"],
        "Total TTC": donnees["total_ttc"],
    }

    if fichier_excel.exists():
        df = pd.read_excel(fichier_excel)
        df = pd.concat([df, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
    else:
        df = pd.DataFrame([nouvelle_ligne])

    # On garantit l'ordre des colonnes attendu.
    df = df.reindex(columns=COLONNES_EXCEL)
    df.to_excel(fichier_excel, index=False)
    logger.info("Ligne ajoutée au tableau de TVA : %s", fichier_excel)


# ---------------------------------------------------------------------------
# 5. ORCHESTRATION
# ---------------------------------------------------------------------------

def traiter_facture(chemin_pdf: Path, email_info: dict, config: dict,
                    clients: list[str]) -> bool:
    """
    Traite UNE facture PDF de bout en bout (analyse -> classement -> Excel).
    Retourne True en cas de succès, False sinon (le PDF reste alors dans
    le dossier temporaire pour vérification manuelle).
    """
    try:
        donnees = analyser_facture(chemin_pdf, clients, email_info["expediteur"])
        destination = classer_facture(chemin_pdf, donnees, config["dossier_base"])
        ajouter_ligne_excel(donnees, destination.name, config["fichier_excel"])
        logger.info(
            "✅ SUCCÈS | %s | client=%s | fournisseur=%s | TTC=%.2f €",
            destination.name, donnees["client"], donnees["fournisseur"],
            donnees["total_ttc"],
        )
        return True
    except Exception as erreur:  # noqa: BLE001 - on isole chaque facture
        logger.error(
            "❌ ÉCHEC  | %s | %s (le PDF reste dans %s pour traitement manuel)",
            chemin_pdf.name, erreur, config["dossier_temp"],
        )
        return False


def main() -> int:
    """Point d'entrée : relève les emails puis traite chaque facture."""
    configurer_logging()
    logger.info("=" * 70)
    logger.info("Démarrage du traitement automatique des factures.")

    try:
        config = charger_configuration()
    except RuntimeError as erreur:
        logger.error("%s", erreur)
        return 1

    # Préparation des dossiers de travail.
    config["dossier_temp"].mkdir(parents=True, exist_ok=True)
    config["dossier_base"].mkdir(parents=True, exist_ok=True)

    clients = charger_clients(config["fichier_clients"])

    try:
        connexion = connecter_imap(config)
    except (imaplib.IMAP4.error, OSError) as erreur:
        logger.error("Connexion IMAP impossible : %s", erreur)
        return 1

    succes, echecs = 0, 0
    try:
        emails_factures = relever_emails_factures(connexion, config)

        for email_info in emails_factures:
            tout_traite = True
            for chemin_pdf in email_info["pdfs"]:
                if traiter_facture(chemin_pdf, email_info, config, clients):
                    succes += 1
                else:
                    echecs += 1
                    tout_traite = False

            # L'email n'est marqué lu que si TOUTES ses factures sont passées,
            # afin qu'un nouveau lancement puisse retenter les échecs.
            if tout_traite:
                marquer_comme_lu(connexion, email_info["id"])
    finally:
        try:
            connexion.close()
            connexion.logout()
        except imaplib.IMAP4.error:
            pass

    logger.info("Traitement terminé : %d succès, %d échec(s).", succes, echecs)
    logger.info("=" * 70)
    return 0 if echecs == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
