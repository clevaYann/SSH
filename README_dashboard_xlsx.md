# Classeur Excel du tableau de bord d'occupation des salles

`build_dashboard_xlsx.py` reproduit le tableau de bord HTML `08_dashboard_complet_<horodatage>.html` dans un classeur Excel
`08_dashboard_complet_<horodatage>.xlsx`, à partir des seuls CSV du dossier de résultats.

## Lancer

```powershell
.venv\Scripts\python build_dashboard_xlsx.py resultats_20261019_100302_987089
# autre horodatage : changer simplement le dossier ; sortie : <dossier>\08_dashboard_complet_<horodatage>.xlsx
.venv\Scripts\python build_dashboard_xlsx.py <dossier> --sortie C:\chemin\classeur.xlsx
```

Prérequis : Python 3.10+, `pandas`, `xlsxwriter` (`pip install pandas xlsxwriter`). La construction prend environ 7 minutes
(2,9 millions de cellules et 1,7 million de formules dans les tableaux de séances). À l'ouverture, Excel recalcule le classeur
(environ 1 à 2 minutes) ; le résultat est lisible dès que « Calcul » disparaît de la barre d'état.

Contrôle (nécessite Excel) : ouvrir le classeur, l'enregistrer sous un autre nom (pour disposer des valeurs recalculées), puis

```powershell
.venv\Scripts\python controle_dashboard_xlsx.py <dossier> <classeur_recalcule.xlsx>
```

Le script compare les indicateurs du classeur à ceux du HTML et écrit `rapport_controle_dashboard_xlsx.md` (et `.csv`).

## Structure du classeur

| Feuille | Origine |
|---|---|
| Conclusion LET, Synthèse V2, Amphis, Réaffectation globale, À vérifier, Usages sans effectif, Référence, H1a, H1b, H1c, H2 (lissage), H3b (lissage), Surdimensionnement | Une feuille par onglet du dashboard, même ordre et mêmes noms. Contiennent uniquement des formules (SOMME.SI.ENS, NB.SI.ENS, MOYENNE.SI.ENS, INDEX/EQUIV, AGREGAT) et des graphiques natifs. |
| Sources | CSV utilisés : nombre de lignes (CSV et classeur), date de génération, feuilles alimentées. |
| Listes | Listes déroulantes, jours pédagogiques (déduits des heures/taux du CSV 28), catalogue des salles lu par formule. |
| Calcul | Minutes d'occupation par salle et par tranche horaire (SOMME.SI.ENS sur les séances). |
| Données_xx | Import brut des CSV en tableaux structurés `t_xx` (filtres). Les CSV 19 à 22 (un par situation) sont empilés avec une colonne SITUATION. `Données_07` ne garde que les 13 colonnes utilisées sur 67. `Données_Report` met à plat les colonnes `<hypothèse>_STATUT/_CODE/_DATE/_DEBUT` du CSV 07. |

Filtres : dans chaque vue de cas, les cellules jaunes (Période, Ville, Site, Bâtiment) sont des listes déroulantes. Dans « Amphis »,
tapez `x` dans la colonne Sélection pour calculer la moyenne des amphis choisis. Dans « Réaffectation globale », choisissez la situation.

Charte : vert `#17675d` (OK), ambre `#b1781b` (à vérifier), rouge `#a6352f` (bloquant), fond `#eef2f1`. Mise en forme conditionnelle sur les
verdicts, les conflits restants et les dépassements de capacité.

## Limites connues (écarts avec le HTML)

- **Pas de liste détaillée des chevauchements** par séance : seul le CSV 13 (par salle et par semaine) existe. La situation de référence
  n'est donc pas détaillée.
- **Grille horaire et heatmap** : recalculées à partir du CSV 07 (taux moyen sur tous les jours de l'année) ; le HTML moyenne cinq taux
  journaliers par jour de semaine, d'où de très faibles différences.
- **Période** : le sélecteur Année / 1er sem. / 2nd sem. / Examens / Hors examens agit sur les taux et heures par salle (CSV 28).
  L'occupation hebdomadaire, la heatmap jour × heure et la grille horaire restent sur l'année complète.
- **Filtre Bâtiment/Ville/Site** : agit sur les KPI, classements, répartitions, hebdomadaire et heatmap ; le détail par salle reste complet (colonne Filtre).
- Les noms de tables sont de la forme `t_28` (et non `T28`, qui est une adresse de cellule pour Excel).
- Excel 365 non requis : aucune fonction matricielle dynamique (pas de FILTRE/TRIER/RECHERCHEX) ; compatible Excel 2013+. Aucune macro VBA.
