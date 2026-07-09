#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
application.py
==============

Application de bureau du traitement automatique des factures.

Interface graphique (Tkinter, inclus avec Python) au-dessus du moteur
traitement_factures.py :

    - bouton « Relever les factures » : lance un cycle complet
      (IMAP -> extraction -> classement -> tableau Excel) sans figer
      la fenêtre (le travail tourne dans un thread) ;
    - historique des factures traitées (lecture du tableau Excel) ;
    - compteur des factures à vérifier manuellement (dossier temporaire) ;
    - journal en direct en bas de fenêtre ;
    - raccourcis : ouvrir le tableau Excel, le dossier des factures,
      le dossier « à vérifier », la liste des clients.

Lancement :
    python application.py        (ou double-clic sur Lancer_Factures.bat)
"""

import logging
import os
import platform
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import pandas as pd

# L'application doit travailler dans son propre dossier (fichiers .env,
# clients.txt, tableau Excel...), quel que soit l'endroit d'où on la lance.
DOSSIER_PROJET = Path(__file__).resolve().parent
os.chdir(DOSSIER_PROJET)

import traitement_factures as moteur  # noqa: E402


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
# Fenêtre principale
# ---------------------------------------------------------------------------

class ApplicationFactures(tk.Tk):
    """Fenêtre principale de l'application."""

    def __init__(self):
        super().__init__()
        self.title("Traitement des factures")
        self.geometry("1080x680")
        self.minsize(860, 560)

        self.file_messages: queue.Queue = queue.Queue()
        self.traitement_en_cours = False
        self.config_moteur: dict | None = None

        self._construire_interface()
        self._brancher_journal()
        self._charger_configuration()
        self.rafraichir_historique()
        self.rafraichir_a_verifier()
        self.after(150, self._vider_file_messages)

    # --- Construction de l'interface ---------------------------------------

    def _construire_interface(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Action.TButton", font=("", 11, "bold"), padding=(14, 8))
        style.configure("Info.TLabel", foreground="#555555")

        # Bandeau du haut : bouton principal + infos de configuration.
        bandeau = ttk.Frame(self, padding=(12, 10))
        bandeau.pack(fill="x")

        self.bouton_relever = ttk.Button(
            bandeau, text="📥  Relever les factures", style="Action.TButton",
            command=self.lancer_traitement,
        )
        self.bouton_relever.pack(side="left")

        self.label_boite = ttk.Label(bandeau, text="", style="Info.TLabel")
        self.label_boite.pack(side="left", padx=16)

        ttk.Button(bandeau, text="📊 Tableau TVA",
                   command=self.ouvrir_excel).pack(side="right", padx=3)
        ttk.Button(bandeau, text="📁 Factures classées",
                   command=self.ouvrir_dossier_factures).pack(side="right", padx=3)
        ttk.Button(bandeau, text="⚠️ À vérifier",
                   command=self.ouvrir_dossier_temp).pack(side="right", padx=3)
        ttk.Button(bandeau, text="👥 Clients",
                   command=self.ouvrir_clients).pack(side="right", padx=3)

        # Historique des factures traitées (contenu du tableau Excel).
        cadre_historique = ttk.LabelFrame(
            self, text=" Factures traitées (les plus récentes en premier) ",
            padding=(8, 6),
        )
        cadre_historique.pack(fill="both", expand=True, padx=12, pady=(2, 6))

        colonnes = ("date_facture", "client", "fournisseur", "ht",
                    "tva20", "tva10", "tva55", "ttc", "fichier")
        entetes = ("Date facture", "Client", "Fournisseur", "Total HT",
                   "TVA 20%", "TVA 10%", "TVA 5.5%", "Total TTC", "Fichier")
        largeurs = (95, 130, 170, 90, 80, 80, 80, 95, 240)

        self.tableau = ttk.Treeview(cadre_historique, columns=colonnes,
                                    show="headings", height=12)
        for colonne, entete, largeur in zip(colonnes, entetes, largeurs):
            self.tableau.heading(colonne, text=entete)
            alignement = "e" if colonne in ("ht", "tva20", "tva10", "tva55", "ttc") else "w"
            self.tableau.column(colonne, width=largeur, anchor=alignement, stretch=True)

        barre = ttk.Scrollbar(cadre_historique, orient="vertical",
                              command=self.tableau.yview)
        self.tableau.configure(yscrollcommand=barre.set)
        self.tableau.pack(side="left", fill="both", expand=True)
        barre.pack(side="right", fill="y")

        # Bandeau d'état : synthèse + factures à vérifier.
        etat = ttk.Frame(self, padding=(12, 0))
        etat.pack(fill="x")
        self.label_synthese = ttk.Label(etat, text="", style="Info.TLabel")
        self.label_synthese.pack(side="left")
        self.label_a_verifier = ttk.Label(etat, text="", foreground="#b3261e")
        self.label_a_verifier.pack(side="right")

        # Journal en direct.
        cadre_journal = ttk.LabelFrame(self, text=" Journal ", padding=(8, 4))
        cadre_journal.pack(fill="both", padx=12, pady=(4, 10))
        self.zone_journal = tk.Text(cadre_journal, height=9, state="disabled",
                                    font=("Consolas", 9), wrap="none")
        self.zone_journal.tag_configure("erreur", foreground="#b3261e")
        self.zone_journal.tag_configure("alerte", foreground="#8a6d00")
        barre_journal = ttk.Scrollbar(cadre_journal, orient="vertical",
                                      command=self.zone_journal.yview)
        self.zone_journal.configure(yscrollcommand=barre_journal.set)
        self.zone_journal.pack(side="left", fill="both", expand=True)
        barre_journal.pack(side="right", fill="y")

    def _brancher_journal(self) -> None:
        """Route les logs du moteur vers la zone de journal de la fenêtre."""
        handler = JournalVersFile(self.file_messages)
        logging.getLogger("factures").setLevel(logging.INFO)
        logging.getLogger("factures").addHandler(handler)
        # Trace également dans le fichier de log habituel.
        fichier = logging.FileHandler("traitement_factures.log", encoding="utf-8")
        fichier.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S"))
        logging.getLogger("factures").addHandler(fichier)

    # --- Configuration -------------------------------------------------------

    def _charger_configuration(self) -> None:
        try:
            self.config_moteur = moteur.charger_configuration()
            self.label_boite.config(
                text=f"Boîte relevée : {self.config_moteur['imap_user']}")
        except RuntimeError as erreur:
            self.label_boite.config(text="⚠ Configuration incomplète (.env)")
            self._journal(logging.ERROR, str(erreur))
            messagebox.showwarning(
                "Configuration",
                f"{erreur}\n\nL'application fonctionne, mais la relève des "
                "emails sera impossible tant que le fichier .env n'est pas rempli.",
            )

    # --- Cycle de traitement -------------------------------------------------

    def lancer_traitement(self) -> None:
        """Démarre un cycle de relève dans un thread (l'UI reste réactive)."""
        if self.traitement_en_cours:
            return
        if self.config_moteur is None:
            self._charger_configuration()
            if self.config_moteur is None:
                return

        self.traitement_en_cours = True
        self.bouton_relever.config(state="disabled",
                                   text="⏳  Relève en cours...")
        threading.Thread(target=self._travail_traitement, daemon=True).start()

    def _travail_traitement(self) -> None:
        """Corps du thread de traitement (ne touche jamais l'UI directement)."""
        try:
            succes, echecs = moteur.executer_traitement(self.config_moteur)
            bilan = f"Cycle terminé : {succes} facture(s) traitée(s), {echecs} échec(s)."
            niveau = logging.WARNING if echecs else logging.INFO
        except Exception as erreur:  # noqa: BLE001 - tout doit remonter à l'UI
            bilan = f"Relève impossible : {erreur}"
            niveau = logging.ERROR
        self.file_messages.put((niveau, bilan))
        self.after(0, self._fin_traitement)

    def _fin_traitement(self) -> None:
        self.traitement_en_cours = False
        self.bouton_relever.config(state="normal", text="📥  Relever les factures")
        self.rafraichir_historique()
        self.rafraichir_a_verifier()

    # --- Rafraîchissement des vues -------------------------------------------

    def rafraichir_historique(self) -> None:
        """Recharge le tableau depuis le fichier Excel global."""
        self.tableau.delete(*self.tableau.get_children())
        fichier_excel = Path(os.getenv("FICHIER_EXCEL", "./tableau_tva_global.xlsx"))
        if not fichier_excel.exists():
            self.label_synthese.config(text="Aucune facture traitée pour le moment.")
            return

        try:
            df = pd.read_excel(fichier_excel)
        except Exception as erreur:  # noqa: BLE001 - fichier ouvert dans Excel ?
            self._journal(logging.WARNING,
                          f"Impossible de lire le tableau Excel : {erreur}")
            return

        for _, ligne in df.iloc[::-1].head(500).iterrows():
            self.tableau.insert("", "end", values=(
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

        total_ttc = pd.to_numeric(df.get("Total TTC"), errors="coerce").sum()
        self.label_synthese.config(
            text=f"{len(df)} facture(s) au total — TTC cumulé : {_euros(total_ttc)} €")

    def rafraichir_a_verifier(self) -> None:
        """Compte les PDF restés dans le dossier temporaire (échecs)."""
        dossier = Path(os.getenv("DOSSIER_TEMP", "./_temp_factures"))
        nombre = len(list(dossier.glob("*.pdf"))) if dossier.exists() else 0
        self.label_a_verifier.config(
            text=f"⚠ {nombre} facture(s) à vérifier manuellement" if nombre else "")

    # --- Journal --------------------------------------------------------------

    def _journal(self, niveau: int, message: str) -> None:
        self.file_messages.put((niveau, message))

    def _vider_file_messages(self) -> None:
        """Déverse la file des logs dans la zone de journal (thread principal)."""
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

    # --- Raccourcis d'ouverture ------------------------------------------------

    def ouvrir_excel(self) -> None:
        _ouvrir(Path(os.getenv("FICHIER_EXCEL", "./tableau_tva_global.xlsx")))

    def ouvrir_dossier_factures(self) -> None:
        _ouvrir(Path(os.getenv("DOSSIER_BASE", "./Factures_Clients")))

    def ouvrir_dossier_temp(self) -> None:
        _ouvrir(Path(os.getenv("DOSSIER_TEMP", "./_temp_factures")))

    def ouvrir_clients(self) -> None:
        _ouvrir(Path(os.getenv("FICHIER_CLIENTS", "./clients.txt")))


# ---------------------------------------------------------------------------
# Utilitaires
# ---------------------------------------------------------------------------

def _euros(valeur) -> str:
    """Format monétaire français : 16896.0 -> '16 896,00'."""
    try:
        if valeur is None or pd.isna(valeur):
            return ""
        return f"{float(valeur):,.2f}".replace(",", " ").replace(".", ",")
    except (TypeError, ValueError):
        return ""


def _ouvrir(chemin: Path) -> None:
    """Ouvre un fichier/dossier avec l'application par défaut du système."""
    chemin.parent.mkdir(parents=True, exist_ok=True)
    if chemin.suffix == "" :
        chemin.mkdir(parents=True, exist_ok=True)
    elif not chemin.exists():
        messagebox.showinfo("Information", f"{chemin.name} n'existe pas encore.")
        return
    systeme = platform.system()
    if systeme == "Windows":
        os.startfile(chemin)  # type: ignore[attr-defined]
    elif systeme == "Darwin":
        subprocess.Popen(["open", str(chemin)])
    else:
        subprocess.Popen(["xdg-open", str(chemin)])


def main() -> int:
    application = ApplicationFactures()
    application.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
