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

### 4 bis. (Recommandé) Adresse email dédiée aux factures

Recevoir les factures sur une adresse dédiée fiabilise le tri : plus de faux
positifs, et vos fournisseurs ont une adresse unique à retenir. Trois options,
de la plus simple à la plus professionnelle :

**Option A — Alias Gmail « +factures » (gratuit, immédiat)**
Gmail livre tout email envoyé à `votreadresse+factures@gmail.com` dans votre
boîte habituelle. Rien à créer côté compte :

1. Communiquez `votreadresse+factures@gmail.com` à vos fournisseurs.
2. Dans Gmail → ⚙️ → *Voir tous les paramètres* → *Filtres et adresses
   bloquées* → *Créer un filtre* : champ **À** = `votreadresse+factures@gmail.com`
   → *Appliquer le libellé* **Factures** (et éventuellement *Ne pas afficher
   dans la boîte de réception* pour ne pas polluer votre boîte principale).
3. Dans `.env` :
   ```
   FILTRE_DESTINATAIRE=votreadresse+factures@gmail.com
   IMAP_FOLDER=Factures        # si vous avez créé le filtre + libellé
   ```

**Option B — Compte Gmail séparé** (ex : `factures.moncentre@gmail.com`) :
créez le compte sur gmail.com, activez la validation en deux étapes, générez
un mot de passe d'application, et mettez ces identifiants dans le `.env`.
`IMAP_FOLDER=INBOX` suffit alors, tout ce qui arrive est une facture.

**Option C — Adresse sur votre propre domaine** (ex :
`factures@moncentre.fr`) : à créer chez votre hébergeur (OVH, Gandi,
Google Workspace...). C'est l'option la plus professionnelle ; renseignez
ensuite `IMAP_HOST` / `IMAP_USER` / `IMAP_PASSWORD` du fournisseur dans `.env`.

> `FILTRE_DESTINATAIRE` restreint la relève aux emails **adressés à** cette
> adresse (en-tête `To`). Combiné au libellé Gmail (`IMAP_FOLDER=Factures`),
> le tri est doublement sécurisé.

### 5. Déclarer vos clients

Éditez `clients.txt` : **un nom de client par ligne** (tel qu'il apparaît sur
les factures — raison sociale de préférence). Le script cherche ces noms dans
le texte du PDF pour savoir à qui appartient chaque facture. Une facture dont
le client n'est pas reconnu est classée dans `Factures_Clients/_A_CLASSER/`.

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

## Limites connues

- Les **PDF scannés** (images sans couche texte) ne sont pas lus : ils sont
  signalés en échec et conservés dans `_temp_factures/`. Un OCR
  (ex. `ocrmypdf` + Tesseract) peut être ajouté si besoin.
- L'extraction par expressions régulières couvre les mises en page de factures
  françaises courantes (« Total HT », « TVA 20 % », « Net à payer »...) ; une
  facture au format très atypique peut nécessiter d'ajuster les motifs dans
  `extraire_montants()` / `extraire_date_facture()`.
- Le fichier Excel ne doit pas être **ouvert dans Excel** pendant l'exécution
  du script (verrouillage du fichier).
