#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
application.py
==============

Application de bureau du traitement automatique des factures.

Interface graphique complète (Tkinter, inclus avec Python) au-dessus du
moteur traitement_factures.py :

    - relève des factures en un clic ou automatique à intervalle régulier ;
    - import manuel de PDF locaux (sans passer par l'email) ;
    - onglet Historique : recherche, double-clic pour ouvrir une facture ;
    - onglet À vérifier : factures en échec, réanalyse, saisie manuelle
      assistée (pré-remplie par l'analyse) ;
    - onglet Synthèse TVA : totaux par mois (HT, TVA par taux, TTC) ;
    - onglet Journal : suivi en direct de chaque étape ;
    - écran de réglages intégré (écrit le fichier .env, test de connexion).

Lancement :
    python application.py        (ou double-clic sur Lancer_Factures.bat)
"""

import imaplib
import logging
import os
import platform
import queue
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from dotenv import dotenv_values

# L'application doit travailler dans son propre dossier (fichiers .env,
# clients.txt, tableau Excel...), quel que soit l'endroit d'où on la lance.
# En version packagée (PyInstaller), __file__ pointe vers un dossier
# temporaire : on prend alors le dossier de l'exécutable.
if getattr(sys, "frozen", False):
    DOSSIER_PROJET = Path(sys.executable).resolve().parent
    # En exécutable fenêtré (sans console), les flux standards sont
    # inexistants : on les remplace pour que bibliothèques et
    # sous-processus ne plantent jamais dessus.
    for _nom_flux in ("stdin", "stdout", "stderr"):
        if getattr(sys, _nom_flux) is None:
            setattr(sys, _nom_flux,
                    open(os.devnull, "r" if _nom_flux == "stdin" else "w"))
else:
    DOSSIER_PROJET = Path(__file__).resolve().parent
os.chdir(DOSSIER_PROJET)

import traitement_factures as moteur  # noqa: E402

TITRE = "Traitement des factures"

# Palette minimaliste : fond neutre, surfaces blanches, un seul accent bleu.
COULEURS = {
    "fond": "#f5f6f8",          # arrière-plan général
    "surface": "#ffffff",       # cartes, tableaux, champs
    "bordure": "#e3e6ea",       # séparations discrètes
    "entete": "#eef1f5",        # en-têtes de tableaux
    "texte": "#1f2430",         # texte principal
    "texte_doux": "#6a7280",    # texte secondaire
    "accent": "#3565d8",        # bleu d'action
    "accent_fonce": "#2a51ad",  # bleu d'action (survol/clic)
    "accent_clair": "#e8eefb",  # sélection, surlignages
    "raye": "#f7f8fa",          # lignes paires des tableaux
    "ok": "#1e7a44",
    "alerte": "#8a6d00",
    "erreur": "#b3261e",
}

# Champs de l'écran de réglages : (clé .env, libellé, valeur par défaut, secret)
CHAMPS_REGLAGES = [
    ("IMAP_HOST", "Serveur IMAP", "imap.gmail.com", False),
    ("IMAP_PORT", "Port IMAP", "993", False),
    ("IMAP_USER", "Adresse email", "mtgsud.compta@gmail.com", False),
    ("IMAP_PASSWORD", "Mot de passe d'application", "", True),
    ("IMAP_FOLDER", "Boîte à relever", "INBOX", False),
    ("MOTS_CLES_OBJET", "Mots-clés de l'objet", "facture,invoice", False),
    ("DOSSIER_BASE", "Dossier des factures classées", "./Factures_Clients", False),
    ("DOSSIER_TEMP", "Dossier temporaire / à vérifier", "./_temp_factures", False),
    ("FICHIER_EXCEL", "Tableau Excel de TVA", "./tableau_tva_global.xlsx", False),
    ("FICHIER_CLIENTS", "Liste des clients", "./clients.txt", False),
]


# Contenu de l'onglet Tutoriel : liste de (style, texte).
TUTORIEL = [
    ("titre", "Guide d'utilisation"),
    ("doux", "Tout ce qu'il faut savoir pour traiter vos factures, de la "
             "première configuration au tableau de TVA."),

    ("titre", "1. Première configuration (une seule fois)"),
    ("etape", "①  Ouvrez ⚙️ Réglages (l'écran s'ouvre tout seul au premier "
              "lancement)."),
    ("etape", "②  Le mot de passe demandé n'est PAS celui de la boîte Gmail : "
              "c'est un « mot de passe d'application » de 16 caractères."),
    ("etape", "③  Pour le créer : connectez-vous au compte Google de la boîte "
              "factures → Sécurité → activez la Validation en deux étapes → "
              "Mots de passe des applications → créez-en un (nom libre)."),
    ("etape", "④  Collez le code dans le champ, cliquez « Tester la "
              "connexion » : un message vert confirme que tout fonctionne."),
    ("etape", "⑤  Cliquez « Enregistrer ». C'est terminé."),

    ("titre", "2. Relever les factures"),
    ("normal", "Cliquez sur le bouton bleu 📥 Relever les factures : "
               "l'application se connecte à la boîte email, télécharge les "
               "PDF des emails non lus, lit chaque facture (montants, TVA, "
               "date, client, fournisseur), la range dans le dossier du "
               "client et remplit le tableau Excel."),
    ("puce", "•  Un email n'est marqué « lu » que si toutes ses factures ont "
             "été traitées : en cas d'échec, il sera retenté au prochain "
             "passage."),
    ("puce", "•  Cochez « Relève automatique » pour que l'application relève "
             "la boîte toute seule à intervalle régulier tant qu'elle est "
             "ouverte."),
    ("puce", "•  📄 Traiter des PDF... permet d'importer des factures depuis "
             "votre ordinateur, sans passer par l'email (l'original n'est "
             "jamais déplacé)."),

    ("titre", "3. L'onglet « Factures traitées »"),
    ("puce", "•  Tapez dans le champ 🔍 pour filtrer par client, fournisseur "
             "ou nom de fichier."),
    ("puce", "•  Double-cliquez sur une ligne pour ouvrir le PDF classé."),
    ("puce", "•  Les factures sont rangées dans Factures_Clients\\<Client>\\ "
             "et renommées : AAAA-MM-JJ_Fournisseur_MontantTTC.pdf."),

    ("titre", "4. L'onglet « À vérifier » : les échecs"),
    ("normal", "Quand une facture ne peut pas être lue automatiquement "
               "(scan de mauvaise qualité, format inhabituel), elle atterrit "
               "ici au lieu d'être classée. Rien n'est perdu :"),
    ("etape", "①  Sélectionnez la facture, cliquez « 👁 Ouvrir le PDF » pour "
              "la lire."),
    ("etape", "②  Cliquez « ✏️ Saisie manuelle... » : le formulaire s'ouvre "
              "déjà pré-rempli avec ce que l'analyse a pu extraire."),
    ("etape", "③  Complétez ou corrigez les montants en lisant le PDF, puis "
              "« Valider et classer » : la facture rejoint les autres."),
    ("puce", "•  « 🔄 Réanalyser » retente la lecture automatique (utile "
             "après installation de Tesseract, par exemple)."),

    ("titre", "5. L'onglet « Synthèse TVA »"),
    ("normal", "Les totaux par mois (HT, TVA 20 %, 10 %, 5,5 %, TTC), "
               "calculés d'après la date de facture : votre base pour la "
               "déclaration de TVA. La ligne TOTAL cumule tout le tableau. "
               "Pour le détail complet, « 📊 Ouvrir le tableau Excel »."),

    ("titre", "6. La liste des clients"),
    ("normal", "L'application reconnaît le client d'une facture en cherchant "
               "son nom dans le texte du PDF. La liste se trouve dans le "
               "fichier clients.txt (un client par ligne). Si un client "
               "apparaît sous plusieurs formes, listez-les après un « = » :"),
    ("code", "ICG 40 = ICG40, I.C.G. 40, SARL ICG 40"),
    ("normal", "Toutes les variantes seront classées dans le même dossier. "
               "Une facture dont le client n'est pas reconnu va dans le "
               "dossier _A_CLASSER."),

    ("titre", "7. Les factures scannées (OCR)"),
    ("normal", "Les factures photographiées ou scannées n'ont pas de texte "
               "lisible : l'application utilise alors Tesseract OCR, à "
               "installer une fois sur l'ordinateur :"),
    ("puce", "•  Téléchargez l'installateur : "
             "github.com/UB-Mannheim/tesseract/wiki"),
    ("puce", "•  Pendant l'installation, à l'étape « Additional language "
             "data », cochez IMPÉRATIVEMENT « French »."),
    ("puce", "•  C'est tout : l'application le détecte automatiquement — la "
             "barre du bas doit afficher « OCR : ✓ Tesseract prêt "
             "(français) » en vert."),
    ("normal", "Si la barre du bas affiche « ⚠ pack français manquant » en "
               "orange : Tesseract est installé mais sans le français. Deux "
               "solutions :"),
    ("puce", "•  Relancez l'installateur de Tesseract et cochez « French » ; "
             "ou"),
    ("puce", "•  Téléchargez le fichier fra.traineddata sur "
             "github.com/tesseract-ocr/tessdata_fast et copiez-le dans "
             "C:\\Program Files\\Tesseract-OCR\\tessdata, puis relancez "
             "l'application."),

    ("titre", "8. Problèmes fréquents"),
    ("puce", "•  « Relève impossible » → vérifiez le mot de passe "
             "d'application dans ⚙️ Réglages (bouton Tester la connexion) et "
             "votre connexion internet."),
    ("puce", "•  « Impossible de lire le tableau Excel » → fermez le fichier "
             "dans Excel pendant que l'application travaille."),
    ("puce", "•  Une facture scannée échoue toujours → scan de qualité "
             "insuffisante : utilisez la saisie manuelle (2 minutes)."),
    ("puce", "•  Le journal (onglet 📜) et le fichier "
             "traitement_factures.log gardent la trace de tout."),
    ("normal", ""),
]


def config_locale() -> dict:
    """Chemins de travail (sans IMAP), pour les opérations locales."""
    return moteur.charger_configuration(exiger_imap=False)


def _euros(valeur) -> str:
    """Format monétaire français : 16896.0 -> '16 896,00'."""
    try:
        if valeur is None or pd.isna(valeur):
            return ""
        return f"{float(valeur):,.2f}".replace(",", " ").replace(".", ",")
    except (TypeError, ValueError):
        return ""


def _centrer_sur(fenetre: tk.Toplevel, parent: tk.Misc) -> None:
    """Positionne un dialogue au-dessus de sa fenêtre parente."""
    fenetre.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - fenetre.winfo_reqwidth()) // 2
    y = parent.winfo_rooty() + 60
    fenetre.geometry(f"+{max(x, 0)}+{max(y, 0)}")


def _ouvrir(chemin: Path) -> None:
    """Ouvre un fichier/dossier avec l'application par défaut du système."""
    if chemin.suffix == "":
        chemin.mkdir(parents=True, exist_ok=True)
    elif not chemin.exists():
        messagebox.showinfo(TITRE, f"{chemin.name} n'existe pas encore.")
        return
    systeme = platform.system()
    if systeme == "Windows":
        os.startfile(chemin)  # type: ignore[attr-defined]
    elif systeme == "Darwin":
        subprocess.Popen(["open", str(chemin)])
    else:
        subprocess.Popen(["xdg-open", str(chemin)])


# ---------------------------------------------------------------------------
# Journalisation vers l'interface
# ---------------------------------------------------------------------------

class JournalVersFile(logging.Handler):
    """Handler logging qui pousse chaque message dans une file lue par l'UI.

    Tkinter n'accepte les mises à jour que depuis le thread principal :
    le thread de traitement écrit ici, la fenêtre vide la file via after().
    """

    def __init__(self, file_messages: queue.Queue):
        super().__init__()
        self.file_messages = file_messages
        self.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))

    def emit(self, record: logging.LogRecord) -> None:
        self.file_messages.put((record.levelno, self.format(record)))


# ---------------------------------------------------------------------------
# Écran de réglages (.env)
# ---------------------------------------------------------------------------

class DialogueReglages(tk.Toplevel):
    """Édition du fichier .env depuis l'interface + test de connexion."""

    def __init__(self, parent: "ApplicationFactures"):
        super().__init__(parent)
        self.parent = parent
        self.title("Réglages")
        self.configure(bg=COULEURS["fond"])
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.saisies: dict[str, tk.Entry] = {}
        self.resultat_test: queue.Queue = queue.Queue()

        valeurs = dotenv_values(".env") if Path(".env").exists() else {}

        corps = ttk.Frame(self, padding=16)
        corps.pack(fill="both", expand=True)

        ttk.Label(corps, text="Connexion à la boîte email",
                  font=("", 10, "bold")).grid(column=0, row=0, columnspan=2,
                                              sticky="w", pady=(0, 6))
        ligne = 1
        for cle, libelle, defaut, secret in CHAMPS_REGLAGES:
            if cle == "MOTS_CLES_OBJET":
                ttk.Separator(corps).grid(column=0, row=ligne, columnspan=2,
                                          sticky="ew", pady=8)
                ligne += 1
                ttk.Label(corps, text="Traitement et classement",
                          font=("", 10, "bold")).grid(column=0, row=ligne,
                                                      columnspan=2, sticky="w",
                                                      pady=(0, 6))
                ligne += 1
            ttk.Label(corps, text=libelle + " :").grid(column=0, row=ligne,
                                                       sticky="w", pady=3)
            champ = ttk.Entry(corps, width=38, show="•" if secret else "")
            champ.insert(0, valeurs.get(cle) or os.getenv(cle) or defaut)
            champ.grid(column=1, row=ligne, sticky="ew", padx=(10, 0), pady=3)
            self.saisies[cle] = champ
            ligne += 1

        # Afficher/masquer le mot de passe.
        self.voir_mdp = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            corps, text="Afficher le mot de passe", variable=self.voir_mdp,
            command=lambda: self.saisies["IMAP_PASSWORD"].config(
                show="" if self.voir_mdp.get() else "•"),
        ).grid(column=1, row=ligne, sticky="w", padx=(10, 0))
        ligne += 1

        ttk.Label(
            corps, foreground="#555555", wraplength=420, justify="left",
            text=("Le mot de passe est un « mot de passe d'application » "
                  "Google (16 caractères), à créer dans : compte Google → "
                  "Sécurité → Validation en deux étapes → Mots de passe des "
                  "applications."),
        ).grid(column=0, row=ligne, columnspan=2, sticky="w", pady=(8, 4))
        ligne += 1

        boutons = ttk.Frame(corps)
        boutons.grid(column=0, row=ligne, columnspan=2, sticky="ew", pady=(12, 0))
        self.bouton_test = ttk.Button(boutons, text="Tester la connexion",
                                      command=self.tester_connexion)
        self.bouton_test.pack(side="left")
        self.label_test = ttk.Label(boutons, text="")
        self.label_test.pack(side="left", padx=10)
        ttk.Button(boutons, text="Enregistrer",
                   command=self.enregistrer).pack(side="right")
        ttk.Button(boutons, text="Annuler",
                   command=self.destroy).pack(side="right", padx=6)

        self.bind("<Return>", lambda _e: self.enregistrer())
        self.saisies["IMAP_PASSWORD"].focus_set()
        _centrer_sur(self, parent)

    # -- Test de connexion (dans un thread, résultat rapatrié par after) ----

    def tester_connexion(self) -> None:
        hote = self.saisies["IMAP_HOST"].get().strip()
        port = self.saisies["IMAP_PORT"].get().strip() or "993"
        utilisateur = self.saisies["IMAP_USER"].get().strip()
        mot_de_passe = self.saisies["IMAP_PASSWORD"].get().strip()
        boite = self.saisies["IMAP_FOLDER"].get().strip() or "INBOX"
        if not (hote and utilisateur and mot_de_passe):
            self.label_test.config(text="Renseignez serveur, email et mot de passe.",
                                   foreground=COULEURS["erreur"])
            return

        self.bouton_test.config(state="disabled")
        self.label_test.config(text="Connexion en cours...", foreground="#555555")

        def travail():
            try:
                connexion = imaplib.IMAP4_SSL(hote, int(port))
                connexion.login(utilisateur, mot_de_passe)
                statut, donnees = connexion.select(boite, readonly=True)
                connexion.logout()
                if statut == "OK":
                    self.resultat_test.put((True, f"Connexion réussie — "
                                            f"{donnees[0].decode()} message(s) dans {boite}."))
                else:
                    self.resultat_test.put((False, f"Boîte « {boite} » introuvable."))
            except Exception as erreur:  # noqa: BLE001
                self.resultat_test.put((False, f"Échec : {erreur}"))

        threading.Thread(target=travail, daemon=True).start()
        self.after(200, self._verifier_resultat_test)

    def _verifier_resultat_test(self) -> None:
        try:
            reussi, message = self.resultat_test.get_nowait()
        except queue.Empty:
            self.after(200, self._verifier_resultat_test)
            return
        self.bouton_test.config(state="normal")
        self.label_test.config(text=message,
                               foreground=COULEURS["ok"] if reussi
                               else COULEURS["erreur"])

    # -- Enregistrement -------------------------------------------------------

    def enregistrer(self) -> None:
        lignes = ["# Configuration du traitement des factures",
                  "# Fichier généré par l'application — ne pas partager.", ""]
        for cle, _libelle, defaut, _secret in CHAMPS_REGLAGES:
            valeur = self.saisies[cle].get().strip() or defaut
            lignes.append(f"{cle}={valeur}")
        Path(".env").write_text("\n".join(lignes) + "\n", encoding="utf-8")
        self.parent.recharger_configuration()
        self.destroy()


# ---------------------------------------------------------------------------
# Saisie manuelle d'une facture en échec
# ---------------------------------------------------------------------------

class DialogueSaisie(tk.Toplevel):
    """Saisie assistée d'une facture que l'analyse n'a pas su lire.

    Les champs sont pré-remplis avec ce que l'analyse a pu extraire ;
    l'utilisateur complète/corrige puis valide : la facture est classée
    et ajoutée au tableau Excel comme une facture traitée normalement.
    """

    def __init__(self, parent: "ApplicationFactures", chemin_pdf: Path,
                 pre_rempli: dict):
        super().__init__(parent)
        self.parent = parent
        self.chemin_pdf = chemin_pdf
        self.title(f"Saisie manuelle — {chemin_pdf.name}")
        self.configure(bg=COULEURS["fond"])
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        corps = ttk.Frame(self, padding=16)
        corps.pack(fill="both", expand=True)

        ttk.Button(corps, text="👁  Ouvrir le PDF pour lecture",
                   command=lambda: _ouvrir(chemin_pdf)).grid(
            column=0, row=0, columnspan=2, sticky="w", pady=(0, 10))

        clients = sorted({canonique for _v, canonique in
                          moteur.charger_clients(config_locale()["fichier_clients"])})
        clients.append(moteur.DOSSIER_NON_CLASSE)

        self.champs: dict[str, tk.Widget] = {}
        definitions = [
            ("date_facture", "Date de facture (AAAA-MM-JJ)"),
            ("client", "Client"),
            ("fournisseur", "Fournisseur"),
            ("total_ht", "Total HT (€)"),
            ("tva_20", "TVA 20 % (€)"),
            ("tva_10", "TVA 10 % (€)"),
            ("tva_5_5", "TVA 5,5 % (€)"),
            ("total_ttc", "Total TTC (€)"),
        ]
        for ligne, (cle, libelle) in enumerate(definitions, start=1):
            ttk.Label(corps, text=libelle + " :").grid(column=0, row=ligne,
                                                       sticky="w", pady=3)
            if cle == "client":
                champ = ttk.Combobox(corps, values=clients, width=33)
            else:
                champ = ttk.Entry(corps, width=36)
            valeur = pre_rempli.get(cle)
            if valeur is not None:
                champ.insert(0, str(valeur))
            champ.grid(column=1, row=ligne, sticky="ew", padx=(10, 0), pady=3)
            self.champs[cle] = champ

        self.label_erreur = ttk.Label(corps, text="",
                                      foreground=COULEURS["erreur"],
                                      wraplength=380)
        self.label_erreur.grid(column=0, row=len(definitions) + 1,
                               columnspan=2, sticky="w", pady=(6, 0))

        boutons = ttk.Frame(corps)
        boutons.grid(column=0, row=len(definitions) + 2, columnspan=2,
                     sticky="ew", pady=(10, 0))
        ttk.Button(boutons, text="Valider et classer",
                   command=self.valider).pack(side="right")
        ttk.Button(boutons, text="Annuler",
                   command=self.destroy).pack(side="right", padx=6)
        _centrer_sur(self, parent)

    def valider(self) -> None:
        def montant(cle):
            texte = self.champs[cle].get().strip()
            return moteur.convertir_montant(texte) if texte else None

        date_facture = self.champs["date_facture"].get().strip()
        try:
            datetime.strptime(date_facture, "%Y-%m-%d")
        except ValueError:
            self.label_erreur.config(text="Date invalide : format attendu AAAA-MM-JJ.")
            return

        total_ttc = montant("total_ttc")
        if total_ttc is None:
            self.label_erreur.config(text="Le Total TTC est obligatoire.")
            return

        donnees = {
            "client": self.champs["client"].get().strip() or moteur.DOSSIER_NON_CLASSE,
            "fournisseur": self.champs["fournisseur"].get().strip() or "Fournisseur_inconnu",
            "date_facture": date_facture,
            "total_ht": montant("total_ht"),
            "tva_20": montant("tva_20"),
            "tva_10": montant("tva_10"),
            "tva_5_5": montant("tva_5_5"),
            "total_ttc": total_ttc,
        }

        try:
            config = config_locale()
            destination = moteur.classer_facture(self.chemin_pdf, donnees,
                                                 config["dossier_base"])
            moteur.ajouter_ligne_excel(donnees, destination.name,
                                       config["fichier_excel"])
        except Exception as erreur:  # noqa: BLE001
            self.label_erreur.config(text=f"Enregistrement impossible : {erreur}")
            return

        self.parent.rafraichir_tout()
        self.destroy()


# ---------------------------------------------------------------------------
# Fenêtre principale
# ---------------------------------------------------------------------------

class ApplicationFactures(tk.Tk):
    """Fenêtre principale de l'application."""

    def __init__(self):
        super().__init__()
        self.title(TITRE)
        self.geometry("1120x720")
        self.minsize(900, 600)

        self.file_messages: queue.Queue = queue.Queue()
        self.traitement_en_cours = False
        self.config_moteur: dict | None = None
        self.derniere_releve: float | None = None

        self._configurer_style()
        self._construire_barre_haut()
        self._construire_onglets()
        self._construire_barre_etat()
        self._brancher_journal()

        self.recharger_configuration(silencieux=True)
        self.rafraichir_tout()
        self.after(150, self._vider_file_messages)
        self.after(15_000, self._tic_releve_auto)

        # État de la chaîne OCR, affiché d'emblée : c'est LE point qui
        # détermine si les factures scannées pourront être lues.
        statut_ocr, message_ocr = moteur.diagnostic_ocr()
        self._journal(logging.INFO if statut_ocr == "ok" else logging.WARNING,
                      message_ocr)
        libelles = {
            "ok": ("OCR : ✓ Tesseract prêt (français)", COULEURS["ok"]),
            "partiel": ("OCR : ⚠ pack français manquant — lecture dégradée "
                        "(voir Tutoriel §7)", COULEURS["alerte"]),
            "absent": ("OCR : ✗ Tesseract non installé (factures scannées "
                       "illisibles — voir Tutoriel §7)", COULEURS["erreur"]),
        }
        texte_ocr, couleur_ocr = libelles[statut_ocr]
        self.label_ocr.config(text=texte_ocr, foreground=couleur_ocr)

        # Premier lancement : ouvrir directement les réglages.
        if not Path(".env").exists():
            self.after(300, self.ouvrir_reglages)

    # --- Style ---------------------------------------------------------------

    def _configurer_style(self) -> None:
        c = COULEURS
        self.configure(bg=c["fond"])
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        # Base : tout repose sur le fond neutre.
        style.configure(".", background=c["fond"], foreground=c["texte"],
                        bordercolor=c["bordure"], focuscolor=c["accent_clair"])
        style.configure("TFrame", background=c["fond"])
        style.configure("TLabel", background=c["fond"], foreground=c["texte"])
        style.configure("Info.TLabel", foreground=c["texte_doux"])
        style.configure("TCheckbutton", background=c["fond"])
        style.map("TCheckbutton", background=[("active", c["fond"])])
        style.configure("TSeparator", background=c["bordure"])

        # Boutons : sobres par défaut, accent plein pour l'action principale.
        style.configure("TButton", background=c["surface"], padding=(10, 6),
                        relief="flat", bordercolor=c["bordure"])
        style.map("TButton",
                  background=[("pressed", c["accent_clair"]),
                              ("active", c["accent_clair"])])
        style.configure("Action.TButton", font=("", 11, "bold"),
                        padding=(18, 10), background=c["accent"],
                        foreground="#ffffff", bordercolor=c["accent"])
        style.map("Action.TButton",
                  background=[("disabled", "#a8b6d4"),
                              ("pressed", c["accent_fonce"]),
                              ("active", c["accent_fonce"])],
                  foreground=[("disabled", "#f0f2f7")])

        # Champs de saisie.
        style.configure("TEntry", fieldbackground=c["surface"],
                        bordercolor=c["bordure"], padding=4)
        style.configure("TCombobox", fieldbackground=c["surface"],
                        background=c["surface"], padding=4)

        # Tableaux.
        style.configure("Treeview", rowheight=26, background=c["surface"],
                        fieldbackground=c["surface"], foreground=c["texte"],
                        bordercolor=c["bordure"], relief="flat")
        style.configure("Treeview.Heading", font=("", 9, "bold"),
                        background=c["entete"], foreground=c["texte"],
                        relief="flat", padding=(6, 6))
        style.map("Treeview.Heading", background=[("active", c["entete"])])
        style.map("Treeview",
                  background=[("selected", c["accent_clair"])],
                  foreground=[("selected", c["texte"])])

        # Onglets : plats, l'onglet actif ressort en blanc.
        style.configure("TNotebook", background=c["fond"], borderwidth=0,
                        tabmargins=(8, 6, 8, 0))
        style.configure("TNotebook.Tab", padding=(16, 8),
                        background=c["fond"], foreground=c["texte_doux"],
                        borderwidth=0, font=("", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", c["surface"])],
                  foreground=[("selected", c["texte"])],
                  expand=[("selected", (0, 0, 0, 1))])

        # Barres de défilement discrètes.
        style.configure("Vertical.TScrollbar", background=c["fond"],
                        troughcolor=c["fond"], bordercolor=c["fond"],
                        arrowcolor=c["texte_doux"])

    # --- Barre du haut ---------------------------------------------------------

    def _construire_barre_haut(self) -> None:
        bandeau = ttk.Frame(self, padding=(12, 10, 12, 6))
        bandeau.pack(fill="x")

        self.bouton_relever = ttk.Button(
            bandeau, text="📥  Relever les factures", style="Action.TButton",
            command=self.lancer_traitement,
        )
        self.bouton_relever.pack(side="left")

        ttk.Button(bandeau, text="📄 Traiter des PDF...",
                   command=self.traiter_pdfs_locaux).pack(side="left", padx=(8, 0))

        # Relève automatique.
        cadre_auto = ttk.Frame(bandeau)
        cadre_auto.pack(side="left", padx=18)
        self.auto_active = tk.BooleanVar(value=False)
        ttk.Checkbutton(cadre_auto, text="Relève automatique toutes les",
                        variable=self.auto_active).pack(side="left")
        self.auto_minutes = tk.Spinbox(cadre_auto, from_=5, to=240, width=4,
                                       justify="right", relief="flat",
                                       bg=COULEURS["surface"],
                                       highlightthickness=1,
                                       highlightbackground=COULEURS["bordure"],
                                       buttonbackground=COULEURS["surface"])
        self.auto_minutes.delete(0, "end")
        self.auto_minutes.insert(0, "15")
        self.auto_minutes.pack(side="left", padx=4)
        ttk.Label(cadre_auto, text="min").pack(side="left")

        ttk.Button(bandeau, text="⚙️ Réglages",
                   command=self.ouvrir_reglages).pack(side="right")
        ttk.Button(bandeau, text="📊 Tableau Excel",
                   command=lambda: _ouvrir(config_locale()["fichier_excel"])
                   ).pack(side="right", padx=6)

    # --- Onglets ---------------------------------------------------------------

    def _construire_onglets(self) -> None:
        self.onglets = ttk.Notebook(self)
        self.onglets.pack(fill="both", expand=True, padx=12, pady=(4, 4))
        self._onglet_historique()
        self._onglet_a_verifier()
        self._onglet_synthese()
        self._onglet_journal()
        self._onglet_tutoriel()

    def _onglet_historique(self) -> None:
        cadre = ttk.Frame(self.onglets, padding=8)
        self.onglets.add(cadre, text="  📋 Factures traitées  ")

        barre = ttk.Frame(cadre)
        barre.pack(fill="x", pady=(0, 6))
        ttk.Label(barre, text="🔍 Rechercher :").pack(side="left")
        self.recherche = tk.StringVar()
        self.recherche.trace_add("write",
                                 lambda *_a: self.rafraichir_historique())
        ttk.Entry(barre, textvariable=self.recherche, width=32).pack(
            side="left", padx=6)
        ttk.Label(barre, text="(double-clic sur une ligne : ouvrir la facture)",
                  style="Info.TLabel").pack(side="left", padx=4)
        self.label_synthese = ttk.Label(barre, text="", style="Info.TLabel")
        self.label_synthese.pack(side="right")

        colonnes = ("date_facture", "client", "fournisseur", "ht",
                    "tva20", "tva10", "tva55", "ttc", "fichier")
        entetes = ("Date facture", "Client", "Fournisseur", "Total HT",
                   "TVA 20%", "TVA 10%", "TVA 5.5%", "Total TTC", "Fichier")
        largeurs = (92, 120, 165, 88, 78, 78, 78, 92, 250)

        self.tableau = ttk.Treeview(cadre, columns=colonnes, show="headings")
        for colonne, entete, largeur in zip(colonnes, entetes, largeurs):
            self.tableau.heading(colonne, text=entete)
            alignement = "e" if colonne in ("ht", "tva20", "tva10", "tva55", "ttc") else "w"
            self.tableau.column(colonne, width=largeur, anchor=alignement)
        self.tableau.tag_configure("paire", background=COULEURS["raye"])
        self.tableau.bind("<Double-1>", self._ouvrir_facture_selectionnee)

        barre_v = ttk.Scrollbar(cadre, orient="vertical",
                                command=self.tableau.yview)
        self.tableau.configure(yscrollcommand=barre_v.set)
        self.tableau.pack(side="left", fill="both", expand=True)
        barre_v.pack(side="right", fill="y")

    def _onglet_a_verifier(self) -> None:
        cadre = ttk.Frame(self.onglets, padding=8)
        self.onglets.add(cadre, text="  ⚠ À vérifier  ")

        ttk.Label(cadre, style="Info.TLabel", wraplength=900, justify="left",
                  text=("Les factures que l'analyse automatique n'a pas su lire "
                        "(scan illisible, format inhabituel) restent ici. "
                        "Ouvrez le PDF, puis « Saisie manuelle » pour compléter "
                        "les montants : la facture sera classée normalement."),
                  ).pack(fill="x", pady=(0, 6))

        boutons = ttk.Frame(cadre)
        boutons.pack(fill="x", pady=(0, 6))
        ttk.Button(boutons, text="👁 Ouvrir le PDF",
                   command=self._ouvrir_pdf_a_verifier).pack(side="left")
        ttk.Button(boutons, text="✏️ Saisie manuelle...",
                   command=self._saisie_manuelle).pack(side="left", padx=6)
        ttk.Button(boutons, text="🔄 Réanalyser",
                   command=self._reanalyser).pack(side="left")
        ttk.Button(boutons, text="🗑 Supprimer",
                   command=self._supprimer_a_verifier).pack(side="left", padx=6)
        ttk.Button(boutons, text="📁 Ouvrir le dossier",
                   command=lambda: _ouvrir(config_locale()["dossier_temp"])
                   ).pack(side="right")

        self.tableau_verifier = ttk.Treeview(
            cadre, columns=("fichier", "taille", "recu"), show="headings")
        for colonne, entete, largeur in (("fichier", "Fichier PDF", 520),
                                         ("taille", "Taille", 90),
                                         ("recu", "Reçu le", 150)):
            self.tableau_verifier.heading(colonne, text=entete)
            self.tableau_verifier.column(colonne, width=largeur, anchor="w")
        self.tableau_verifier.bind("<Double-1>",
                                   lambda _e: self._ouvrir_pdf_a_verifier())
        barre_v = ttk.Scrollbar(cadre, orient="vertical",
                                command=self.tableau_verifier.yview)
        self.tableau_verifier.configure(yscrollcommand=barre_v.set)
        self.tableau_verifier.pack(side="left", fill="both", expand=True)
        barre_v.pack(side="right", fill="y")

    def _onglet_synthese(self) -> None:
        cadre = ttk.Frame(self.onglets, padding=8)
        self.onglets.add(cadre, text="  📊 Synthèse TVA  ")

        ttk.Label(cadre, style="Info.TLabel",
                  text="Totaux par mois (d'après la date de facture) — "
                       "la ligne TOTAL cumule l'ensemble du tableau.",
                  ).pack(fill="x", pady=(0, 6))

        colonnes = ("mois", "nb", "ht", "tva20", "tva10", "tva55",
                    "tva_totale", "ttc")
        entetes = ("Mois", "Factures", "Total HT", "TVA 20%", "TVA 10%",
                   "TVA 5.5%", "TVA totale", "Total TTC")
        self.tableau_synthese = ttk.Treeview(cadre, columns=colonnes,
                                             show="headings")
        for colonne, entete in zip(colonnes, entetes):
            self.tableau_synthese.heading(colonne, text=entete)
            self.tableau_synthese.column(
                colonne, width=110, anchor="e" if colonne != "mois" else "w")
        self.tableau_synthese.tag_configure("total", font=("", 9, "bold"),
                                            background=COULEURS["accent_clair"])
        barre_v = ttk.Scrollbar(cadre, orient="vertical",
                                command=self.tableau_synthese.yview)
        self.tableau_synthese.configure(yscrollcommand=barre_v.set)
        self.tableau_synthese.pack(side="left", fill="both", expand=True)
        barre_v.pack(side="right", fill="y")

    def _onglet_journal(self) -> None:
        cadre = ttk.Frame(self.onglets, padding=8)
        self.onglets.add(cadre, text="  📜 Journal  ")
        self.zone_journal = tk.Text(cadre, state="disabled", wrap="none",
                                    font=("Consolas", 9), height=8,
                                    bg=COULEURS["surface"], fg=COULEURS["texte"],
                                    relief="flat", highlightthickness=1,
                                    highlightbackground=COULEURS["bordure"],
                                    padx=8, pady=6)
        self.zone_journal.tag_configure("erreur", foreground=COULEURS["erreur"])
        self.zone_journal.tag_configure("alerte", foreground=COULEURS["alerte"])
        barre_v = ttk.Scrollbar(cadre, orient="vertical",
                                command=self.zone_journal.yview)
        self.zone_journal.configure(yscrollcommand=barre_v.set)
        self.zone_journal.pack(side="left", fill="both", expand=True)
        barre_v.pack(side="right", fill="y")

    def _onglet_tutoriel(self) -> None:
        cadre = ttk.Frame(self.onglets, padding=0)
        self.onglets.add(cadre, text="  ❓ Tutoriel  ")

        texte = tk.Text(cadre, wrap="word", relief="flat", height=8,
                        bg=COULEURS["surface"], fg=COULEURS["texte"],
                        padx=28, pady=20, spacing1=4, spacing3=6,
                        highlightthickness=0, cursor="arrow")
        barre_v = ttk.Scrollbar(cadre, orient="vertical", command=texte.yview)
        texte.configure(yscrollcommand=barre_v.set)
        texte.pack(side="left", fill="both", expand=True)
        barre_v.pack(side="right", fill="y")

        texte.tag_configure("titre", font=("", 14, "bold"),
                            foreground=COULEURS["accent"],
                            spacing1=18, spacing3=8)
        texte.tag_configure("soustitre", font=("", 11, "bold"), spacing1=10)
        texte.tag_configure("normal", font=("", 10))
        texte.tag_configure("doux", font=("", 10),
                            foreground=COULEURS["texte_doux"])
        texte.tag_configure("puce", font=("", 10), lmargin1=18, lmargin2=32)
        texte.tag_configure("etape", font=("", 10), lmargin1=18, lmargin2=36)
        texte.tag_configure("code", font=("Consolas", 9),
                            background=COULEURS["entete"], lmargin1=18)

        for etiquette, contenu in TUTORIEL:
            texte.insert("end", contenu + "\n", etiquette)
        texte.configure(state="disabled")

    # --- Barre d'état ------------------------------------------------------------

    def _construire_barre_etat(self) -> None:
        barre = ttk.Frame(self, padding=(12, 4, 12, 8))
        barre.pack(fill="x")
        self.label_boite = ttk.Label(barre, text="", style="Info.TLabel")
        self.label_boite.pack(side="left")
        self.label_ocr = ttk.Label(barre, text="", style="Info.TLabel")
        self.label_ocr.pack(side="left", padx=24)
        self.label_etat = ttk.Label(barre, text="Prêt.", style="Info.TLabel")
        self.label_etat.pack(side="right")

    # --- Configuration -------------------------------------------------------------

    def recharger_configuration(self, silencieux: bool = False) -> None:
        try:
            self.config_moteur = moteur.charger_configuration()
            self.label_boite.config(
                text=f"Boîte relevée : {self.config_moteur['imap_user']}")
        except RuntimeError as erreur:
            self.config_moteur = None
            self.label_boite.config(
                text="⚠ Configuration incomplète — ouvrez ⚙️ Réglages")
            if not silencieux:
                self._journal(logging.ERROR, str(erreur))

    def ouvrir_reglages(self) -> None:
        DialogueReglages(self)

    # --- Cycle de traitement ---------------------------------------------------------

    def lancer_traitement(self) -> None:
        """Démarre un cycle de relève dans un thread (l'UI reste réactive)."""
        if self.traitement_en_cours:
            return
        if self.config_moteur is None:
            self.recharger_configuration()
            if self.config_moteur is None:
                self.ouvrir_reglages()
                return

        self._debut_travail("⏳  Relève en cours...")
        threading.Thread(target=self._travail_releve, daemon=True).start()

    def _travail_releve(self) -> None:
        """Corps du thread de relève (ne touche jamais l'UI directement)."""
        try:
            succes, echecs = moteur.executer_traitement(self.config_moteur)
            bilan = f"Cycle terminé : {succes} facture(s) traitée(s), {echecs} échec(s)."
            niveau = logging.WARNING if echecs else logging.INFO
        except Exception as erreur:  # noqa: BLE001 - tout doit remonter à l'UI
            bilan = f"Relève impossible : {erreur}"
            niveau = logging.ERROR
        self.derniere_releve = time.monotonic()
        self.file_messages.put((niveau, bilan))
        self.after(0, self._fin_travail, bilan)

    def traiter_pdfs_locaux(self) -> None:
        """Traite des PDF choisis sur le disque, sans passer par l'email."""
        fichiers = filedialog.askopenfilenames(
            title="Choisir des factures PDF",
            filetypes=[("Factures PDF", "*.pdf")])
        if not fichiers:
            return
        self._debut_travail("⏳  Traitement des PDF...")
        threading.Thread(target=self._travail_pdfs_locaux,
                         args=(list(fichiers),), daemon=True).start()

    def _travail_pdfs_locaux(self, fichiers: list[str]) -> None:
        config = config_locale()
        config["dossier_temp"].mkdir(parents=True, exist_ok=True)
        config["dossier_base"].mkdir(parents=True, exist_ok=True)
        clients = moteur.charger_clients(config["fichier_clients"])
        info = {"expediteur": "Import manuel"}
        succes = echecs = 0
        for fichier in fichiers:
            # Copie dans le dossier temporaire : l'original n'est jamais touché,
            # et un échec laisse la copie dans « À vérifier » comme pour un email.
            copie = config["dossier_temp"] / Path(fichier).name
            try:
                shutil.copy2(fichier, copie)
            except OSError as erreur:
                self.file_messages.put((logging.ERROR,
                                        f"Copie impossible ({fichier}) : {erreur}"))
                echecs += 1
                continue
            if moteur.traiter_facture(copie, info, config, clients):
                succes += 1
            else:
                echecs += 1
        bilan = f"Import terminé : {succes} facture(s) traitée(s), {echecs} échec(s)."
        self.file_messages.put(
            (logging.WARNING if echecs else logging.INFO, bilan))
        self.after(0, self._fin_travail, bilan)

    def _debut_travail(self, texte: str) -> None:
        self.traitement_en_cours = True
        self.bouton_relever.config(state="disabled", text=texte)
        self.label_etat.config(text=texte.strip("⏳ "))

    def _fin_travail(self, bilan: str) -> None:
        self.traitement_en_cours = False
        self.bouton_relever.config(state="normal",
                                   text="📥  Relever les factures")
        self.label_etat.config(
            text=f"{bilan}  ({datetime.now().strftime('%H:%M')})")
        self.rafraichir_tout()

    # --- Relève automatique -------------------------------------------------------

    def _tic_releve_auto(self) -> None:
        try:
            if self.auto_active.get() and not self.traitement_en_cours:
                try:
                    intervalle = max(5, int(self.auto_minutes.get())) * 60
                except ValueError:
                    intervalle = 15 * 60
                if (self.derniere_releve is None
                        or time.monotonic() - self.derniere_releve >= intervalle):
                    self._journal(logging.INFO, "Relève automatique...")
                    self.lancer_traitement()
        finally:
            self.after(15_000, self._tic_releve_auto)

    # --- Onglet À vérifier : actions ------------------------------------------------

    def _pdf_a_verifier_selectionne(self) -> Path | None:
        selection = self.tableau_verifier.selection()
        if not selection:
            messagebox.showinfo(TITRE, "Sélectionnez d'abord une facture "
                                       "dans la liste.")
            return None
        nom = self.tableau_verifier.item(selection[0], "values")[0]
        return config_locale()["dossier_temp"] / nom

    def _ouvrir_pdf_a_verifier(self) -> None:
        chemin = self._pdf_a_verifier_selectionne()
        if chemin:
            _ouvrir(chemin)

    def _saisie_manuelle(self) -> None:
        chemin = self._pdf_a_verifier_selectionne()
        if not chemin:
            return
        # Pré-remplissage best-effort : on tente l'analyse (peut prendre
        # quelques secondes si OCR) puis on ouvre le formulaire.
        self.label_etat.config(text="Analyse du PDF pour pré-remplissage...")
        self.update_idletasks()
        pre_rempli: dict = {}
        try:
            clients = moteur.charger_clients(config_locale()["fichier_clients"])
            pre_rempli = moteur.analyser_facture(chemin, clients, "")
        except Exception:  # noqa: BLE001 - la saisie reste possible à vide
            try:
                texte = moteur.extraire_texte_pdf(chemin)
                montants = moteur.extraire_montants(texte)
                pre_rempli = {
                    "date_facture": moteur.extraire_date_facture(texte),
                    "fournisseur": moteur.identifier_fournisseur(texte),
                    "total_ht": montants["total_ht"],
                    "tva_20": montants["tva"]["20"],
                    "tva_10": montants["tva"]["10"],
                    "tva_5_5": montants["tva"]["5.5"],
                    "total_ttc": montants["total_ttc"],
                }
            except Exception:  # noqa: BLE001
                pre_rempli = {}
        self.label_etat.config(text="Prêt.")
        DialogueSaisie(self, chemin, pre_rempli)

    def _reanalyser(self) -> None:
        chemin = self._pdf_a_verifier_selectionne()
        if not chemin:
            return
        self._debut_travail("⏳  Réanalyse...")
        threading.Thread(target=self._travail_reanalyse, args=(chemin,),
                         daemon=True).start()

    def _travail_reanalyse(self, chemin: Path) -> None:
        config = config_locale()
        clients = moteur.charger_clients(config["fichier_clients"])
        reussi = moteur.traiter_facture(chemin, {"expediteur": "Reprise manuelle"},
                                        config, clients)
        bilan = ("Réanalyse réussie, facture classée."
                 if reussi else "Réanalyse : échec — utilisez la saisie manuelle.")
        self.after(0, self._fin_travail, bilan)

    def _supprimer_a_verifier(self) -> None:
        chemin = self._pdf_a_verifier_selectionne()
        if not chemin:
            return
        if messagebox.askyesno(TITRE, f"Supprimer définitivement "
                                      f"{chemin.name} ?"):
            chemin.unlink(missing_ok=True)
            self.rafraichir_a_verifier()

    # --- Rafraîchissement des vues ---------------------------------------------------

    def rafraichir_tout(self) -> None:
        self.rafraichir_historique()
        self.rafraichir_a_verifier()
        self.rafraichir_synthese()

    def _lire_excel(self) -> pd.DataFrame | None:
        fichier = config_locale()["fichier_excel"]
        if not fichier.exists():
            return None
        try:
            return pd.read_excel(fichier)
        except Exception as erreur:  # noqa: BLE001 - fichier ouvert dans Excel ?
            self._journal(logging.WARNING,
                          f"Impossible de lire le tableau Excel : {erreur}")
            return None

    def rafraichir_historique(self) -> None:
        self.tableau.delete(*self.tableau.get_children())
        df = self._lire_excel()
        if df is None or df.empty:
            self.label_synthese.config(text="Aucune facture traitée pour le moment.")
            return

        filtre = self.recherche.get().strip().lower()
        if filtre:
            masque = (
                df.get("Client", "").astype(str).str.lower().str.contains(filtre, na=False)
                | df.get("Fournisseur", "").astype(str).str.lower().str.contains(filtre, na=False)
                | df.get("Nom du fichier", "").astype(str).str.lower().str.contains(filtre, na=False)
            )
            df_affiche = df[masque]
        else:
            df_affiche = df

        for indice, (_, ligne) in enumerate(df_affiche.iloc[::-1].head(1000).iterrows()):
            self.tableau.insert("", "end",
                                tags=("paire",) if indice % 2 else (),
                                values=(
                ligne.get("Date de facture", ""),
                ligne.get("Client", ""),
                ligne.get("Fournisseur", ""),
                _euros(ligne.get("Total HT")),
                _euros(ligne.get("TVA 20%")),
                _euros(ligne.get("TVA 10%")),
                _euros(ligne.get("TVA 5.5%")),
                _euros(ligne.get("Total TTC")),
                ligne.get("Nom du fichier", ""),
            ))

        total_ttc = pd.to_numeric(df_affiche.get("Total TTC"),
                                  errors="coerce").sum()
        self.label_synthese.config(
            text=f"{len(df_affiche)} facture(s) — TTC cumulé : "
                 f"{_euros(total_ttc)} €")

    def rafraichir_a_verifier(self) -> None:
        self.tableau_verifier.delete(*self.tableau_verifier.get_children())
        dossier = config_locale()["dossier_temp"]
        fichiers = sorted(dossier.glob("*.pdf")) if dossier.exists() else []
        for pdf in fichiers:
            infos = pdf.stat()
            self.tableau_verifier.insert("", "end", values=(
                pdf.name,
                f"{infos.st_size / 1024:.0f} Ko",
                datetime.fromtimestamp(infos.st_mtime).strftime("%d/%m/%Y %H:%M"),
            ))
        # Met à jour le titre de l'onglet avec le compteur.
        self.onglets.tab(1, text=f"  ⚠ À vérifier ({len(fichiers)})  "
                         if fichiers else "  ⚠ À vérifier  ")

    def rafraichir_synthese(self) -> None:
        self.tableau_synthese.delete(*self.tableau_synthese.get_children())
        df = self._lire_excel()
        if df is None or df.empty:
            return

        df = df.copy()
        df["_mois"] = pd.to_datetime(df["Date de facture"],
                                     errors="coerce").dt.strftime("%Y-%m")
        df["_mois"] = df["_mois"].fillna("(sans date)")
        numeriques = ["Total HT", "TVA 20%", "TVA 10%", "TVA 5.5%", "Total TTC"]
        for colonne in numeriques:
            df[colonne] = pd.to_numeric(df.get(colonne), errors="coerce")

        groupes = df.groupby("_mois", sort=True)
        for mois, groupe in groupes:
            tva_totale = groupe[["TVA 20%", "TVA 10%", "TVA 5.5%"]].sum().sum()
            self.tableau_synthese.insert("", "end", values=(
                mois, len(groupe),
                _euros(groupe["Total HT"].sum()),
                _euros(groupe["TVA 20%"].sum()),
                _euros(groupe["TVA 10%"].sum()),
                _euros(groupe["TVA 5.5%"].sum()),
                _euros(tva_totale),
                _euros(groupe["Total TTC"].sum()),
            ))
        tva_totale = df[["TVA 20%", "TVA 10%", "TVA 5.5%"]].sum().sum()
        self.tableau_synthese.insert("", "end", tags=("total",), values=(
            "TOTAL", len(df),
            _euros(df["Total HT"].sum()),
            _euros(df["TVA 20%"].sum()),
            _euros(df["TVA 10%"].sum()),
            _euros(df["TVA 5.5%"].sum()),
            _euros(tva_totale),
            _euros(df["Total TTC"].sum()),
        ))

    # --- Historique : ouverture d'une facture -----------------------------------------

    def _ouvrir_facture_selectionnee(self, _evenement) -> None:
        selection = self.tableau.selection()
        if not selection:
            return
        valeurs = self.tableau.item(selection[0], "values")
        client, fichier = valeurs[1], valeurs[8]
        if not fichier:
            return
        chemin = (config_locale()["dossier_base"]
                  / moteur.nettoyer_nom_fichier(str(client)) / str(fichier))
        if chemin.exists():
            _ouvrir(chemin)
        else:
            messagebox.showinfo(TITRE, f"Fichier introuvable :\n{chemin}")

    # --- Journal -------------------------------------------------------------------

    def _brancher_journal(self) -> None:
        """Route les logs du moteur vers l'onglet Journal + fichier de log."""
        journal = logging.getLogger("factures")
        journal.setLevel(logging.INFO)
        journal.addHandler(JournalVersFile(self.file_messages))
        fichier = logging.FileHandler("traitement_factures.log", encoding="utf-8")
        fichier.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S"))
        journal.addHandler(fichier)

    def _journal(self, niveau: int, message: str) -> None:
        self.file_messages.put((niveau, message))

    def _vider_file_messages(self) -> None:
        """Déverse la file des logs dans l'onglet Journal (thread principal)."""
        try:
            while True:
                niveau, message = self.file_messages.get_nowait()
                etiquette = ("erreur" if niveau >= logging.ERROR
                             else "alerte" if niveau >= logging.WARNING else "")
                self.zone_journal.configure(state="normal")
                self.zone_journal.insert("end", message + "\n", etiquette)
                self.zone_journal.see("end")
                self.zone_journal.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(150, self._vider_file_messages)


def main() -> int:
    application = ApplicationFactures()
    application.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
