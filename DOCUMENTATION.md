# Analyse d'occupation des salles UPPA : comment ça marche

Ce document explique à quoi sert l'application, ce qu'elle lit, comment elle calcule, comment la lancer et comment lire ses résultats. Il s'adresse aux personnes qui relancent l'analyse ou présentent les résultats (Direction du Patrimoine, SSH).

## 1. À quoi ça sert

L'application répond à la question : **peut-on supprimer les trois amphithéâtres du bâtiment Lettres (LET) en reportant leurs cours sur les cinq amphithéâtres de Droit (DEG), puis sur les grandes salles ?**

Elle part du planning réel 2025-2026 et produit :

- la **situation de référence** : taux d'occupation de chaque salle ;
- **quatre simulations de report** (H1, H2, H3a, H3b) ;
- le repérage des **locaux surdimensionnés** ;
- des **indicateurs** (classements, écarts effectif/capacité, périodes d'examen) ;
- un **tableau de bord HTML** unique, filtrable, ouvrable sans serveur ni connexion.

Le programme est un seul fichier Python : `analyse_occupation_uppa.py`.

## 2. Ce que l'application lit

Tous les fichiers d'entrée sont dans le dossier `BDD/`.

| Fichier | Rôle |
|---|---|
| `CC EXPORT MIASHS V2_2025_2026.csv` | Planning : une ligne par modèle de séance (matière, jour, heures, salle, promotions, semaines, effectif `EFFCALCU`). |
| `EXP_SALLE_2025-2026.csv` | Catalogue des salles : code, nom `Nom (BÂTIMENT-ÉTAGE)`, capacité. |
| `EXP_PROMOTION_2025-2026.csv` | Promotions et leur effectif. Sert à reconnaître les promotions et, en option, à estimer un effectif manquant. |
| `SURFACES GENERAL UPPA 2026_09.xlsx` | Code bâtiment → ville et site, pour les filtres du tableau de bord. |
| `dashboard_template.html` | Gabarit visuel du tableau de bord (onglets Vue d'ensemble, Grille horaire, Analyse hebdomadaire, Synthèse exécutive, Chevauchements). |

`EXP_ETUDIANT_2025-2026.csv` et les PDF (`Méthodologie Occupation UPPA.pdf`, `information_complementaire.pdf`) sont de la documentation ou des données non utilisées par les calculs.

Le programme accepte les encodages UTF-8 et cp1252, et détecte le séparateur tout seul.

## 3. Comment le calcul se déroule

### 3.1 Reconstituer les séances réelles

Une ligne du planning n'est **pas** une séance unique : c'est un modèle qui se répète. La colonne `PERIODE` (par exemple `[3..6,8..13,16]`) liste les semaines ISO où la séance a lieu. Le programme crée donc une séance par semaine listée, au jour de la semaine de `DDEBUT`.

Il applique ensuite ces règles :

1. **Blocages administratifs exclus** : une réservation de 8 h ou plus n'est pas un cours.
2. **Week-ends exclus.**
3. **Jours fériés nationaux exclus** (calculés, y compris Pâques, Ascension, Pentecôte).
4. **Fermetures de l'université exclues** : Toussaint, Noël, hiver, printemps (voir l'option `--fermetures`).
5. **Doublons fusionnés** : plusieurs lignes qui décrivent la même séance (même date, salle, heures, matière, type) deviennent une seule séance. Les promotions sont regroupées.
6. **Contrôles de cohérence** : `DDEBUT` doit appartenir à `PERIODE`, le jour doit correspondre à la date, les heures doivent être lisibles. Sinon le programme s'arrête avec un message clair plutôt que de produire des résultats faux.

**Rattachement aux salles.** Chaque séance est reliée à une salle par `CODE_SAL`. Un code préfixé `<>` (par exemple `<>17105`) est accepté quand le nom de la salle correspond exactement au catalogue. Les séances sans `CODE_SAL` ne sont rattachées à aucune salle.

**Jours pédagogiques.** Ce sont les jours ouvrés où l'export contient au moins une séance, après retrait des fériés et des fermetures. C'est la base de tous les taux.

### 3.2 Calculer le taux d'occupation

```
taux d'une salle = heures occupées entre 08:00 et 18:00 / (jours pédagogiques × 10 h)
```

- Seules les minutes entre 8 h et 18 h comptent.
- Les séances qui se chevauchent dans une même salle sont **conservées** : le taux peut donc dépasser 100 %, ce qui signale une double réservation.
- Les taux sont aussi calculés par semestre (S1 : août à janvier ; S2 : février à juillet), par semaine ISO et par créneau d'une heure.

### 3.3 Types de salles

Le type est déduit du nom dans le catalogue :

- **amphi** : le nom contient « amphi » ;
- **grande salle** : « salle de cours » de 50 à 100 places dans LET ou DEG ;
- **salle de cours** et **autre** pour le reste.

### 3.4 Les scénarios

Les séances à reporter sont toutes celles des trois amphis LET. Dans chaque scénario, les séances déjà programmées dans les salles d'accueil restent inchangées.

| Cas | Salles d'accueil | Règle |
|---|---|---|
| **Référence** | aucune | Situation réelle, rien n'est déplacé. |
| **H1** | 5 amphis DEG | Même jour, même heure. Choix du plus petit amphi libre assez grand. Si aucun n'est libre, la séance est placée quand même dans le plus petit amphi assez grand et **le chevauchement est enregistré**. |
| **H2** | 5 amphis DEG | Chaque séance peut changer de jour ou d'heure. Elle prend le créneau libre le plus proche de l'original. |
| **H3a** | 5 amphis DEG + grandes salles de LET et DEG | Comme H1. |
| **H3b** | 5 amphis DEG + grandes salles de LET et DEG | Comme H2. |

**Contraintes du lissage (H2 et H3b)** :

- la durée de la séance ne change pas ;
- la capacité de la salle est suffisante ;
- la séance reste entre 08:00 et 18:00 ;
- une salle n'accueille jamais deux séances en même temps ;
- une promotion (`NOM_DIP`) n'a jamais deux séances en même temps.

Les séances sont traitées dans l'ordre chronologique. Une séance déplacée peut donc prendre le créneau d'une séance suivante, qui doit alors se déplacer à son tour.

**Statuts d'une séance après report :**

| Statut | Signification |
|---|---|
| `place` | Placée dans un local libre de capacité suffisante. |
| `place_avec_chevauchement` | H1 et H3a : aucun local libre à ce créneau, la séance est placée quand même. |
| `sans_solution` | H2 et H3b : aucun créneau libre n'a été trouvé. |
| `non_couvert` | Aucun local d'accueil n'est assez grand. |
| `effectif_inconnu` | Effectif absent, nul ou contradictoire : la séance n'est pas placée (voir 3.5). |

Le programme vérifie à chaque exécution que **les heures sont conservées** : heures source = heures placées + heures non placées.

### 3.5 Les effectifs manquants

L'effectif de chaque séance est `EFFCALCU`. Il n'est **jamais inventé** : s'il manque, la séance est signalée `effectif_inconnu` et n'est pas reportée. Ses heures disparaissent donc des salles d'accueil dans les simulations.

L'option `--effectifs-estimes` est la seule exception. Pour les séances de la forme « `<Promotion>groupe` », elle retient l'effectif de la promotion parente dans `EXP_PROMOTION` comme **majorant** (borne haute). Cette estimation s'est révélée supérieure ou égale à l'effectif réel dans 96 % des séances testées. Les séances concernées sont marquées « estimé » dans les CSV et dans le dashboard. Les classements et la référence n'utilisent jamais ces estimations.

### 3.6 Locaux surdimensionnés

Une séance de type Cours, CM, TD ou CTD est surdimensionnée si la capacité de sa salle vaut au moins deux fois l'effectif et laisse au moins 20 places vides. Le programme propose, pour chaque série de séances, le plus petit local suffisant libre à toutes les occurrences, de préférence dans le même bâtiment.

### 3.7 Indicateurs

- **Top 5 / Flop 5** des salles les plus et les moins occupées (amphis et grandes salles de LET et DEG).
- **Remplissage moyen** par salle (effectif / capacité) et classement des plus forts écarts.
- **Focus examens** : les mêmes indicateurs sur les seules périodes d'examen (option `--examens`).
- **Chevauchements** par scénario, salle, semaine et période.
- **Séances déplaçables** vers une salle plus adaptée, et créneaux d'amphi libérés.

## 4. Lancer l'application

### 4.1 Prérequis

- Python 3.10 ou plus récent.
- Le module `openpyxl` (lecture du fichier Excel). Il est déjà installé dans `.venv`.

### 4.2 Commandes

Depuis le dossier du projet :

```powershell
# Analyse standard
.venv\Scripts\python analyse_occupation_uppa.py

# Avec estimation des effectifs manquants (à mentionner si vous présentez les résultats)
.venv\Scripts\python analyse_occupation_uppa.py --effectifs-estimes

# Test automatique du programme sur des données synthétiques
.venv\Scripts\python analyse_occupation_uppa.py --self-test
```

Chaque exécution crée un **nouveau dossier** `resultats_AAAAMMJJ_HHMMSS_xxxxxx/` : les anciens résultats ne sont jamais écrasés.

### 4.3 Options

| Option | Rôle | Valeur par défaut |
|---|---|---|
| `--dossier` | Dossier contenant les fichiers d'entrée | `BDD/` |
| `--planning`, `--salles`, `--promotions` | Chemins explicites des fichiers | détection automatique |
| `--surfaces` | Classeur Excel ville/site | `BDD/SURFACES*.xlsx` |
| `--modele-dashboard` | Gabarit du dashboard | `BDD/dashboard_template.html` |
| `--sortie` | Dossier où créer les résultats | dossier du projet |
| `--fermetures` | Fermetures `AAAA-MM-JJ:AAAA-MM-JJ`, séparées par des virgules | vacances 2025-2026 |
| `--examens` | Périodes d'examen, même format | 5-17 janvier et 18 mai-30 juin 2026 |
| `--debut`, `--fin` | Limiter l'analyse à une période (`AAAA-MM-JJ`) | tout l'export |
| `--exclure-semaines` | Semaines ISO à exclure (`43,52,1`) | aucune |
| `--effectifs-estimes` | Estimer les effectifs manquants (voir 3.5) | désactivé |
| `--modele-effectifs` | Écrit la liste des promotions à renseigner puis s'arrête | désactivé |

**Réutiliser pour une autre année** : changez les fichiers dans `BDD/` et passez les nouvelles dates à `--fermetures` et `--examens`. Aucune modification du code n'est nécessaire. Le calcul des jours fériés couvre les années 2025 et 2026 : pour une autre année, adaptez la fonction correspondante.

## 5. Lire les résultats

### 5.1 Le fichier à ouvrir : `08_dashboard_complet_….html`

C'est un fichier unique, sans connexion requise. Sous les onglets, un **panneau d'explication en langage simple** décrit la vue choisie et donne les chiffres de cette exécution.

La synthèse (`00_synthese_….html`) et la page des indicateurs (`15_indicateurs_….html`) ne sont plus des onglets du dashboard : ce sont des fichiers séparés dans le même dossier.

| Onglet | Contenu |
|---|---|
| **Amphis** | Tableau des amphis seuls, trié et filtré (voir 5.4). |
| **Cours sans effectif** | Les cours des amphis de Lettres qui ne peuvent pas être déplacés faute de connaître le nombre d'étudiants : explication, chiffres, liste complète triable et téléchargement du CSV à compléter. |
| **Référence, H1, H2, H3a, H3b** | Tableau de bord complet de chaque cas (vue d'ensemble, grille horaire, analyse hebdomadaire, synthèse, chevauchements) avec filtres ville, site, bâtiment, recherche et taux, et sélecteur de période (voir 5.2). |
| **Surdimensionnement** | Séances accueillies dans un local trop grand, et salle plus adaptée proposée. |

### 5.2 Les périodes d'examen

Dans chaque vue de cas, le sélecteur de période propose **Année**, **1er sem.**, **2nd sem.**, **Examens** et **Hors examens**. Les périodes d'examen par défaut sont **début janvier (5 au 17 janvier 2026)** et **fin mai à juin (18 mai au 30 juin 2026)** ; on les change avec `--examens`.

Le choix s'applique à tous les onglets : taux et heures de chaque salle, grille horaire, graphique hebdomadaire (semaines contenant un jour d'examen), synthèse exécutive (nombre de jours et de créneaux) et liste des chevauchements. La page Amphis a le même sélecteur. « Hors examens » = tous les autres jours pédagogiques, donc heures d'examen + heures hors examens = heures de l'année.

**Attention :** l'export s'arrête au 30 avril 2026. La période de mai-juin ne contient donc aucune donnée, et « Examens » ne couvre aujourd'hui que dix jours de janvier.

### 5.3 Pourquoi le « taux d'occupation moyen » ne change presque pas d'un scénario à l'autre

Ce KPI porte sur toutes les salles du catalogue. Un report **déplace** des heures de LET vers DEG sans en créer ni en supprimer : le total, donc la moyenne, reste quasi identique. Pour mesurer l'effet d'un scénario, regardez :

- le taux des salles d'accueil (onglet Amphis ou Indicateurs) ;
- le nombre de chevauchements ;
- la saturation par salle et par semaine.

### 5.4 L'onglet Amphis

- Une ligne par amphi, avec le taux et les heures dans chaque cas, et l'écart par rapport à la référence.
- Des filtres à cases à cocher (ville, site, bâtiment), une recherche et trois boutons : tous les amphis, LET (supprimés), DEG (accueil).
- Une case à cocher devant chaque amphi. Le bloc **Moyenne générale** calcule, pour les amphis cochés (ou les amphis affichés si aucun n'est coché), le taux moyen, l'écart avec la référence et les heures totales.
- Un clic sur un titre de colonne trie ; un second clic inverse l'ordre.
- Un sélecteur **Année / Examens / Hors examens** recalcule les taux, les heures et la moyenne pour la période choisie.

### 5.5 Les autres fichiers du dossier de résultats

| Fichier | Contenu |
|---|---|
| `00_synthese`, `01_reference`, `02` à `05` (H1 à H3b) | Les mêmes pages que le dashboard, une par fichier. |
| `06_surdimensionnement` (.html et .csv) | Séances surdimensionnées. |
| `07_detail_seances.csv` | **Une ligne par séance** : statut après report dans chaque scénario, nouvelle salle, nouveau créneau, capacité, écart, effectif estimé ou non. |
| `09_occupation_avant_apres.csv` | Heures et taux de chaque salle du périmètre, avant et après chaque scénario, sur l'année et sur les examens. |
| `10_classement_occupation.csv` | Top 5 et flop 5. |
| `11_remplissage_par_salle.csv` | Remplissage moyen et écart moyen par salle. |
| `12_ecarts_effectif_capacite.csv` (et `_examens`) | Chaque séance classée par écart. |
| `13_chevauchements_par_salle_semaine.csv` | Chevauchements par scénario, salle, semaine et période. |
| `17_cours_sans_effectif.csv` (et `.html`) | Liste des cours non déplacés faute d'effectif, avec une colonne vide `EFFECTIF_A_RENSEIGNER` à compléter. |
| `14_anomalies.csv` | **Journal d'anomalies** avec une explication pour chaque type. |
| `15_indicateurs.html`, `16_amphis.html` | Pages Indicateurs et Amphis seules. |

Les CSV utilisent le séparateur `;` et l'encodage UTF-8 avec BOM : ils s'ouvrent directement dans Excel.

## 6. Limites à connaître avant de décider

Ces points sont aussi écrits dans la page « Indicateurs » du dashboard.

1. **Séances sans salle.** Environ 45 % des lignes de l'export n'ont aucun `CODE_SAL`. Elles comptent pour les jours ouverts mais pas dans l'occupation d'une salle. Si certaines se déroulent en réalité dans les amphis LET ou DEG, l'occupation est sous-estimée.
2. **Effectifs manquants.** Environ 300 séances des amphis LET n'ont aucun effectif et ne sont pas reportées. La faisabilité repose sur les autres.
3. **Effectif supérieur à la capacité.** Quelques séances dépassent la capacité de leur salle. Elles sont conservées et listées dans le journal d'anomalies : il faut vérifier la capacité ou l'effectif.
4. **Fenêtre de l'export.** L'export va du 8 septembre 2025 au 30 avril 2026. Il n'y a rien pour mai et juin : les « examens » se limitent à quelques jours de janvier.
5. **Lissage.** C'est un test théorique de disponibilité, pas un emploi du temps validé. Les enseignants ne figurent pas dans l'export : leur disponibilité n'est pas contrôlée. Les plateaux d'examen ne sont pas modélisés.
6. **Rapprochements de bâtiments.** Quelques codes du catalogue sont rattachés à d'autres codes du classeur Excel (`B45→B4B5`, `DAL_R→DAL`, `SCI_R→SCI`, `IPM→IPM1`, `XPL→XLP`, `GTR→RT`, `TEL→Télésite Tarbes`). Les bâtiments `BDL` et `BTB` n'ont pas de correspondance : leurs salles n'ont ni ville ni site.

## 7. Fiabilité : ce qui est contrôlé

- À chaque exécution, un contrôle recompte les séances des amphis LET directement dans le planning brut et compare avec le moteur. Un écart arrête le programme.
- Le programme refuse de continuer si le calendrier est incohérent, si une capacité est dépassée dans un scénario ou si des heures disparaissent.
- Les résultats de la dernière version ont aussi été recalculés par un programme indépendant : séances, taux de toutes les salles, chevauchements, capacités, ville et site, absence de séance un jour fermé. Aucun écart n'a été trouvé.
- `--self-test` vérifie le lecteur de fichiers et les scénarios sur des cas synthétiques.

## 8. En cas de problème

| Message ou symptôme | Cause probable |
|---|---|
| `ARRÊT : … PERIODE illisible` | Format de semaine inattendu dans le planning. Le message donne la valeur fautive. |
| `ARRÊT : Contrôle calendrier échoué` | `DDEBUT` ne correspond pas à `PERIODE` ou au jour indiqué. |
| `Colonnes manquantes dans …` | Le fichier d'entrée n'a pas les colonnes attendues (voir section 2). |
| Filtres ville/site vides | Le classeur `SURFACES*.xlsx` est introuvable ou `openpyxl` n'est pas installé. |
| `Aucun fichier SURFACES trouvé` | Idem : indiquez le chemin avec `--surfaces`. |
| Peu de jours d'examen | La période demandée sort de la fenêtre de l'export (voir limite 4). |
