# 📄 Traitement automatique des factures par email

Script Python local qui relève une boîte mail (IMAP), télécharge les factures
PDF, en extrait les données comptables (client, fournisseur, HT, TVA par taux,
TTC, date), classe chaque facture dans le dossier du client concerné et
alimente un tableau Excel global de TVA.

## Fonctionnement

```
Boîte mail (IMAP)
      │  emails non lus : objet "facture"/"invoice" OU pièce jointe PDF
      ▼
_temp_factures/            ← téléchargement des PDF
      │  extraction : pdfplumber + expressions régulières
      ▼
Factures_Clients/
  └── <Nom_du_client>/     ← créé automatiquement si absent
        └── YYYY-MM-DD_Fournisseur_MontantTTC.pdf
      │
      ▼
tableau_tva_global.xlsx    ← une ligne par facture traitée
```

Colonnes du tableau Excel : Date de traitement, Date de facture, Client,
Fournisseur, Nom du fichier, Total HT, TVA 20%, TVA 10%, TVA 5.5%, Total TTC.

Règles de sécurité intégrées :
- les identifiants IMAP sont lus depuis un fichier `.env` (jamais dans le code) ;
- un email n'est marqué « lu » que si **toutes** ses factures ont été traitées
  avec succès — un échec sera donc retenté au prochain lancement ;
- une facture en échec reste dans `_temp_factures/` pour vérification manuelle ;
- tous les événements sont tracés dans le terminal **et** dans le fichier
  `traitement_factures.log`.

## Installation pas à pas

### 1. Prérequis

- Python **3.10 ou plus** (`python3 --version` pour vérifier).

### 2. Récupérer le projet et créer un environnement virtuel

```bash
git clone <url-du-depot>
cd comptabilit-
python3 -m venv .venv
source .venv/bin/activate        # Windows : .venv\Scripts\activate
```

### 3. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 4. Configurer le fichier `.env`

```bash
cp .env.example .env             # Windows : copy .env.example .env
```

Puis ouvrez `.env` et renseignez au minimum :

| Variable        | Description                                      | Exemple                  |
|-----------------|--------------------------------------------------|--------------------------|
| `IMAP_HOST`     | Serveur IMAP de votre messagerie                 | `imap.gmail.com`         |
| `IMAP_PORT`     | Port IMAP SSL (993 en général)                   | `993`                    |
| `IMAP_USER`     | Votre adresse email                              | `contact@moncentre.fr`   |
| `IMAP_PASSWORD` | Mot de passe (voir note ci-dessous)              | `abcd efgh ijkl mnop`    |

> **Important pour Gmail** : le mot de passe de votre compte ne fonctionnera
> pas. Créez un **mot de passe d'application** : compte Google → Sécurité →
> Validation en deux étapes (à activer) → Mots de passe des applications.
> Même principe pour Outlook/Office 365.

Les autres variables (`DOSSIER_BASE`, `DOSSIER_TEMP`, `FICHIER_EXCEL`,
`MOTS_CLES_OBJET`, ...) sont optionnelles et documentées dans `.env.example`.

### 3 bis. Installer Tesseract (OCR des factures scannées)

En pratique, la majorité des factures fournisseurs sont des **scans** (photos
ou numérisations sans texte). Le script les lit grâce à l'OCR Tesseract, qui
doit être installé sur la machine :

- **Windows** : téléchargez l'installateur sur
  https://github.com/UB-Mannheim/tesseract/wiki — pendant l'installation,
  cochez le pack de langue **French**. Si `tesseract` n'est pas dans le PATH,
  ajoutez dans le `.env` ou en début de script :
  `pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"`
- **macOS** : `brew install tesseract tesseract-lang`
- **Linux (Debian/Ubuntu)** : `sudo apt install tesseract-ocr tesseract-ocr-fra`

Sans Tesseract, le script fonctionne quand même : les PDF scannés sont
simplement signalés en échec et conservés dans `_temp_factures/`.

### 4 bis. La boîte dédiée aux factures : mtgsud.compta@gmail.com

Le traitement utilise une boîte Gmail entièrement dédiée aux factures :
**`mtgsud.compta@gmail.com`**. Tout ce qui y arrive est une facture (ou un
transfert de facture), ce qui rend le tri fiable à 100 %.

Mise en service du compte (une seule fois) :

1. Connectez-vous à `mtgsud.compta@gmail.com` sur https://myaccount.google.com
2. **Sécurité** → activez la **Validation en deux étapes**.
3. Toujours dans Sécurité → **Mots de passe des applications** → créez-en un
   (nom libre, ex. « script factures »). Google affiche un code de
   16 caractères : c'est lui qu'il faut mettre dans `IMAP_PASSWORD` du `.env`.
4. Vérifiez que IMAP est actif : Gmail → ⚙️ → *Voir tous les paramètres* →
   *Transfert et POP/IMAP* → *Activer IMAP* (actif par défaut sur les
   comptes récents).
5. Communiquez l'adresse à vos fournisseurs, ou transférez-y vos factures
   depuis votre boîte principale.

> Astuce : les emails **transférés** (objet « Fwd: ... ») sont traités comme
> les autres — c'est la pièce jointe PDF qui compte, pas l'objet.

Si un jour vous relevez une boîte partagée plutôt que dédiée, la variable
`FILTRE_DESTINATAIRE` permet de ne traiter que les emails adressés à un
alias donné (ex : `adresse+factures@gmail.com`).

### 5. Déclarer vos clients

Éditez `clients.txt` : **un client par ligne**, tel qu'il apparaît sur les
factures. Si un client apparaît sous plusieurs formes, listez-les après un
`=` — toutes les variantes seront classées dans le même dossier :

```
ICG 40 = ICG40, I.C.G. 40, SARL ICG 40
SCI LANGON
```

Le script cherche ces noms dans le texte du PDF pour savoir à qui appartient
chaque facture. Une facture dont le client n'est pas reconnu est classée dans
`Factures_Clients/_A_CLASSER/`.

### 6. Lancer le script

```bash
python traitement_factures.py
```

Exemple de sortie :

```
2026-07-09 09:12:01 | INFO     | Connecté à la boîte 'INBOX'.
2026-07-09 09:12:02 | INFO     | 3 email(s) non lu(s) trouvé(s).
2026-07-09 09:12:03 | INFO     | Pièce jointe téléchargée : facture_edf.pdf
2026-07-09 09:12:04 | INFO     | Facture classée : Factures_Clients/SARL_Dupont_Consulting/2026-06-28_EDF_142.80.pdf
2026-07-09 09:12:04 | INFO     | ✅ SUCCÈS | 2026-06-28_EDF_142.80.pdf | client=SARL Dupont Consulting | fournisseur=EDF | TTC=142.80 €
2026-07-09 09:12:05 | ERROR    | ❌ ÉCHEC  | scan_facture.pdf | Aucun texte extrait (PDF scanné/image ? un OCR serait nécessaire).
2026-07-09 09:12:05 | INFO     | Traitement terminé : 1 succès, 1 échec(s).
```

### 7. (Optionnel) Automatiser le lancement

- **Linux/macOS** — toutes les heures via cron :
  ```
  0 * * * * cd /chemin/vers/comptabilit- && .venv/bin/python traitement_factures.py
  ```
- **Windows** — Planificateur de tâches, action :
  `C:\chemin\.venv\Scripts\python.exe C:\chemin\traitement_factures.py`

## Tests

Le dossier `tests/` contient une suite de non-régression calquée sur des
factures réelles (acomptes négatifs, sous-totaux, colonnes OCR sans libellés,
numéros de TVA intracommunautaire, tickets de caisse...) :

```bash
python tests/test_extraction.py     # ou : python -m pytest tests/
```

## Limites connues

- **PDF scannés** : lus par OCR (Tesseract). Un scan de très mauvaise qualité
  (décimales illisibles, texte trop dégradé) est signalé en échec et conservé
  dans `_temp_factures/` pour saisie manuelle — le script ne devine jamais un
  montant douteux.
- **Paiements partiels** : le script extrait le **Total TTC de la facture**
  (la donnée comptable), pas le montant du virement. Une facture payée en
  plusieurs fois apparaît donc avec son total, ce qui est le comportement
  attendu pour le tableau de TVA.
- L'extraction par expressions régulières couvre les mises en page de factures
  françaises courantes (« Total HT », « TVA 20 % », « Net à payer », « Somme à
  payer »...) ; une facture au format très atypique peut nécessiter d'ajuster
  les motifs dans `extraire_montants()` / `extraire_date_facture()`.
- Le fichier Excel ne doit pas être **ouvert dans Excel** pendant l'exécution
  du script (verrouillage du fichier).
