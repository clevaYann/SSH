#!/usr/bin/env python3
"""Reproduit le tableau de bord HTML d'occupation des salles (dossier resultats_…) dans un classeur Excel.

Principe
- Les données viennent exclusivement des CSV du dossier (jamais du HTML) : chaque CSV est importé tel quel dans un
  tableau Excel structuré (ListObject) d'une feuille « Données_xx ».
- Les feuilles de restitution (une par onglet du dashboard) ne contiennent que des FORMULES (SOMME.SI.ENS,
  NB.SI.ENS, INDEX/EQUIV, AGREGAT…) qui pointent vers ces tableaux : aucune valeur calculée par pandas n'est collée.
- Graphiques Excel natifs alimentés par les plages de formules. Aucune macro VBA.

Usage : python build_dashboard_xlsx.py <dossier_resultats> [--sortie fichier.xlsx]
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import json
import os
import re
import sys
from pathlib import Path

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

# --- Charte graphique du dashboard HTML ------------------------------------------------------
OK, WARN, BAD, BG = "#17675d", "#b1781b", "#a6352f", "#eef2f1"
INK, MUTED, LINE = "#17252a", "#52636b", "#d1dadb"

SITUATIONS = ["Référence", "H1a", "H1b", "H1c", "H2", "H3b"]
VIEW_SHEETS = {"Référence": "Référence", "H1a": "H1a", "H1b": "H1b", "H1c": "H1c", "H2": "H2 (lissage)", "H3b": "H3b (lissage)"}
PERIODS = ["Année", "1er sem.", "2nd sem.", "Examens", "Hors examens"]
HOURS = list(range(8, 18))
DAYS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi"]
BUCKETS = [(0, 5), (5, 10), (10, 15), (15, 20), (20, 25), (25, 30), (30, 40), (40, 10**6)]
ALL = "Toutes"  # valeur « pas de filtre » des listes déroulantes
NOT_MATCH = "<>§§"  # critère SOMME.SI.ENS qui accepte toutes les valeurs, y compris les cellules vides


# =============================================================================================
# 1. Lecture des CSV
# =============================================================================================
def read_csv(path: str) -> pd.DataFrame:
    """Lit un CSV du dossier (séparateur « ; », UTF-8 avec BOM, repli cp1252)."""
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return pd.read_csv(path, sep=";", encoding=enc, dtype=str, keep_default_na=False)
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"Encodage illisible : {path}")


def numeric(df: pd.DataFrame, text_cols: set[str]) -> pd.DataFrame:
    """Convertit en nombres les colonnes entièrement numériques (hors colonnes à garder en texte)."""
    out = df.copy()
    for col in out.columns:
        if col in text_cols:
            continue
        converted = pd.to_numeric(out[col].replace("", pd.NA), errors="coerce")
        if converted.notna().sum() == out[col].replace("", pd.NA).notna().sum() and converted.notna().any():
            out[col] = converted
    return out


class Sources:
    """Inventaire des CSV du dossier (nom court → chemin) et des feuilles qu'ils alimentent."""

    def __init__(self, folder: str):
        self.folder = folder
        self.files = {}
        for path in sorted(glob.glob(os.path.join(folder, "*.csv"))):
            name = os.path.basename(path)
            self.files[name] = path
        stamp = re.search(r"_(\d{8}_\d{6}_\d+)\.csv$", next(iter(self.files)))
        self.stamp = stamp.group(1) if stamp else ""
        self.used: dict[str, dict] = {}

    def find(self, prefix: str, suffix: str = "") -> str:
        exact = [p for n, p in self.files.items() if n == f"{prefix}{suffix}_{self.stamp}.csv"]
        if exact:
            return exact[0]
        hits = [p for n, p in self.files.items() if n.startswith(prefix) and (not suffix or n.endswith(f"{suffix}_{self.stamp}.csv"))]
        if suffix == "" and len(hits) > 1:
            hits = [p for p in hits if re.fullmatch(rf"{re.escape(prefix)}[A-Za-z0-9_]*_{self.stamp}\.csv", os.path.basename(p)) and not re.search(r"_(H\d\w?|reference)_", os.path.basename(p))]
        if len(hits) != 1:
            raise SystemExit(f"CSV introuvable ou ambigu pour {prefix}{suffix} : {hits}")
        return hits[0]

    def register(self, path: str, rows: int, sheets: str, note: str = ""):
        self.used[os.path.basename(path)] = {"rows": rows, "date": dt.datetime.fromtimestamp(os.path.getmtime(path)), "sheets": sheets, "note": note}


# =============================================================================================
# 2. Classeur : formats et utilitaires
# =============================================================================================
class Book:
    def __init__(self, path: str):
        self.wb = xlsxwriter.Workbook(path, {"use_future_functions": True, "nan_inf_to_errors": False})
        self.wb.set_calc_mode("auto")
        f = self.wb.add_format
        base = {"font_name": "Calibri", "font_size": 10, "font_color": INK}
        self.fmt = {
            "base": f(base),
            "title": f({**base, "bold": True, "font_size": 18, "font_color": OK}),
            "sub": f({**base, "font_color": MUTED, "italic": True}),
            "h2": f({**base, "bold": True, "font_size": 12, "font_color": OK, "bottom": 2, "bottom_color": OK}),
            "th": f({**base, "bold": True, "font_color": "#ffffff", "bg_color": OK, "text_wrap": True, "valign": "vcenter", "border": 1, "border_color": LINE}),
            "label": f({**base, "font_color": MUTED, "font_size": 9, "bold": True, "text_wrap": True}),
            "sel": f({**base, "bold": True, "bg_color": "#fff6dd", "border": 1, "border_color": WARN}),
            "kpi_l": f({**base, "font_color": MUTED, "font_size": 9, "bg_color": "#ffffff", "border": 1, "border_color": LINE, "text_wrap": True, "valign": "vcenter", "align": "center"}),
            "kpi_n": f({**base, "bold": True, "font_size": 16, "font_color": OK, "bg_color": "#ffffff", "border": 1, "border_color": LINE, "align": "center", "num_format": "#,##0"}),
            "kpi_h": f({**base, "bold": True, "font_size": 16, "font_color": OK, "bg_color": "#ffffff", "border": 1, "border_color": LINE, "align": "center", "num_format": '#,##0" h"'}),
            "kpi_p": f({**base, "bold": True, "font_size": 16, "font_color": OK, "bg_color": "#ffffff", "border": 1, "border_color": LINE, "align": "center", "num_format": "0.0%"}),
            "kpi_t": f({**base, "bold": True, "font_size": 12, "font_color": OK, "bg_color": "#ffffff", "border": 1, "border_color": LINE, "align": "center", "text_wrap": True}),
            "txt": f({**base, "border": 1, "border_color": LINE, "text_wrap": True, "valign": "top"}),
            "txt_b": f({**base, "border": 1, "border_color": LINE, "text_wrap": True, "valign": "top", "bold": True}),
            "num": f({**base, "border": 1, "border_color": LINE, "num_format": "#,##0"}),
            "num1": f({**base, "border": 1, "border_color": LINE, "num_format": "#,##0.0"}),
            "hrs": f({**base, "border": 1, "border_color": LINE, "num_format": '#,##0" h"'}),
            "hrs1": f({**base, "border": 1, "border_color": LINE, "num_format": '#,##0.0" h"'}),
            "pct": f({**base, "border": 1, "border_color": LINE, "num_format": "0.0%"}),
            "date": f({**base, "num_format": "dd/mm/yyyy"}),
            "note": f({**base, "font_color": MUTED, "text_wrap": True, "valign": "top"}),
            "good": f({"bg_color": "#d9ece8", "font_color": OK, "bold": True}),
            "warn": f({"bg_color": "#f8ebd0", "font_color": WARN, "bold": True}),
            "bad": f({"bg_color": "#f4dcd9", "font_color": BAD, "bold": True}),
        }
        self.sheets: dict[str, xlsxwriter.worksheet.Worksheet] = {}
        self.cache: dict[tuple[str, str], object] = {}

    def sheet(self, name: str, tab: str | None = None, hide_grid: bool = True):
        ws = self.wb.add_worksheet(name)
        ws.set_default_row(15)
        if hide_grid:
            ws.hide_gridlines(2)
        if tab:
            ws.set_tab_color(tab)
        self.sheets[name] = ws
        return ws

    def f(self, ws, row: int, col: int, formula: str, fmt: str = "num"):
        """Écrit une formule (valeur mise en cache si elle est connue d'une exécution précédente)."""
        key = (ws.get_name(), xl_rowcol_to_cell(row, col))
        value = self.cache.get(key, 0)
        ws.write_formula(row, col, formula, self.fmt[fmt] if isinstance(fmt, str) else fmt, value)

    def t(self, ws, row: int, col: int, text: str, fmt: str = "base"):
        ws.write(row, col, text, self.fmt[fmt])


def a1(row: int, col: int, absolute: bool = True) -> str:
    return xl_rowcol_to_cell(row, col, row_abs=absolute, col_abs=absolute)


def rng(row1: int, col1: int, row2: int, col2: int) -> str:
    return f"{a1(row1, col1)}:{a1(row2, col2)}"


def sref(sheet: str, cell: str) -> str:
    return f"'{sheet}'!{cell}"


# =============================================================================================
# 3. Feuilles « Données_xx » : import brut des CSV en tableaux Excel structurés
# =============================================================================================
TEXT_COLS = {"CODE", "code", "CODE_SAL", "ID_SEANCE", "H1a_CODE", "target_code", "replacementCode", "ID", "HDEBUT", "HFIN", "JOUR", "PERIODE"}
DATE_COLS = {"DATE", "date"}


def clean(df: pd.DataFrame, text_cols: set[str] | None = None) -> pd.DataFrame:
    """Prépare un DataFrame pour l'écriture : nombres, dates ; cellules vides conservées vides."""
    keep = TEXT_COLS | (text_cols or set())
    out = numeric(df, keep)
    for col in out.columns:
        if col in DATE_COLS:
            out[col] = pd.to_datetime(out[col], errors="coerce")
    return out


def write_table(bk: Book, sheet: str, table: str, df: pd.DataFrame, calc_cols: list[tuple[str, str, str]] | None = None,
                widths: dict[str, int] | None = None, tab: str = "#9aa7b1", start_row: int = 0, ws=None, style: str = "Table Style Light 9"):
    """Écrit `df` dans un tableau structuré. `calc_cols` = colonnes calculées (nom, formule structurée, format)."""
    own = ws is None
    if own:
        ws = bk.sheet(sheet, tab, hide_grid=False)
    calc_cols = calc_cols or []
    cols = []
    for col in df.columns:
        c = {"header": str(col)}
        if col in DATE_COLS:
            c["format"] = bk.fmt["date"]
        cols.append(c)
    for name, formula, fmtname in calc_cols:
        # Références structurées : [@Col] → table[@[Col]] (forme exigée par le format de fichier Excel)
        formula = re.sub(r"(?<![\w\]])\[@\[?([^\]\[]+)\]?\]", lambda m: f"{table}[@[{m.group(1)}]]", formula)
        c = {"header": name, "formula": formula}
        if fmtname:
            c["format"] = bk.fmt[fmtname]
        cols.append(c)
    data = df.astype(object).where(df.notna(), None).values.tolist()
    n = len(data)
    ws.add_table(start_row, 0, start_row + max(n, 1), len(cols) - 1, {"name": table, "columns": cols, "data": data, "style": style, "autofilter": True})
    for i, col in enumerate(list(df.columns) + [c[0] for c in calc_cols]):
        w = (widths or {}).get(col)
        if w is None:
            w = 12 if col in DATE_COLS else min(40, max(8, len(str(col)) + 2))
        ws.set_column(i, i, w)
    if own:
        ws.freeze_panes(start_row + 1, 0)
    return ws


def stacked(src: Sources, prefix: str) -> pd.DataFrame:
    """Empile les CSV « par situation » (19 à 22) dans un seul tableau avec une colonne SITUATION."""
    frames = []
    for key, label in (("reference", "Référence"), ("H1a", "H1a"), ("H1b", "H1b"), ("H1c", "H1c"), ("H2", "H2"), ("H3b", "H3b")):
        path = src.find(prefix, f"_{key}")
        df = read_csv(path)
        src.register(path, len(df), "", "")
        df.insert(0, "SITUATION", label)
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def hour_formulas() -> list[tuple[str, str, str]]:
    """Minutes de la séance dans chaque tranche d'une heure de 8 h à 18 h."""
    return [(f"M{h:02d}", f"=MAX(0,MIN([@Fin],{(h + 1) * 60})-MAX([@Deb],{h * 60}))", "") for h in HOURS]


def minutes(col: str) -> str:
    """Formule structurée : « HH:MM » → minutes depuis minuit."""
    return f"VALUE(LEFT([@{col}],2))*60+VALUE(MID([@{col}],4,2))"


def build_data_sheets(bk: Book, src: Sources) -> dict:
    """Importe tous les CSV. Renvoie des informations utiles à la suite (listes de valeurs, tailles)."""
    info: dict = {}
    # --- 07 : séances (colonnes utiles seulement ; les 67 colonnes restent dans le CSV) ---
    p07 = src.find("07_detail_seances")
    full07 = read_csv(p07)
    keep = ["ID_SEANCE", "DATE", "JOUR", "HDEBUT", "HFIN", "DUREE_HEURES", "TYPE", "LIBELLE_MAT", "NOM_DIP", "NOM_SAL", "CODE_SAL", "CAPA", "EFFCALCU"]
    SAMPLE = int(os.environ.get("XLSX_ECHANTILLON", "0") or 0)  # mise au point : limite les gros tableaux (jamais en production)
    if SAMPLE:
        full07 = full07.head(SAMPLE)
    t07 = clean(full07[keep])
    if not t07["DATE"].is_monotonic_increasing:
        raise SystemExit("07_detail_seances n'est pas trié par date : le comptage des jours pédagogiques serait faux.")
    src.register(p07, len(full07), "Données_07 ; Données_Report ; Référence ; H1a ; H1b ; H1c ; H2 (lissage) ; H3b (lissage)",
                 f"{len(keep)} colonnes sur {len(full07.columns)} importées (celles utilisées)")
    info["weeks"] = sorted({int(d.isocalendar()[0]) * 100 + int(d.isocalendar()[1]) for d in t07["DATE"].dropna()})
    cat_cols = [
        ("Idx", "=IFERROR(MATCH([@[CODE_SAL]],Cat_Code,0),0)", ""),
        ("Bat", '=IF([@Idx]=0,"",INDEX(Cat_Bat,[@Idx]))', ""),
        ("Ville", '=IF([@Idx]=0,"",INDEX(Cat_Ville,[@Idx]))', ""),
        ("Site", '=IF([@Idx]=0,"",INDEX(Cat_Site,[@Idx]))', ""),
        ("SrcLET", "=IF([@Idx]=0,0,INDEX(Cat_LET,[@Idx]))", ""),
    ]
    calc07 = [
        ("Deb", f"={minutes('HDEBUT')}", ""),
        ("Fin", f"={minutes('HFIN')}+IF({minutes('HFIN')}<[@Deb],1440,0)", ""),
        *hour_formulas(),
        ("H818", "=SUM(t_07[@[M08]:[M17]])/60", "num1"),
        ("Sem", "=YEAR([@DATE]+4-WEEKDAY([@DATE],2))*100+ISOWEEKNUM([@DATE])", ""),
        ("JourFlag", "=IF(ROW()-ROW(t_07[[#Headers],[DATE]])=1,1,IF([@DATE]<>INDEX(t_07[DATE],ROW()-ROW(t_07[[#Headers],[DATE]])-1),1,0))", ""),
        *cat_cols,
    ]
    write_table(bk, "Données_07", "t_07", t07, calc07, {"LIBELLE_MAT": 28, "NOM_DIP": 28, "NOM_SAL": 24})
    info["n07"] = len(t07)

    # --- Report des séances LET : mise à plat des colonnes <hypothèse>_* du CSV 07 (aucun calcul d'indicateur) ---
    rows = []
    for sc in ("H1a", "H1b", "H1c", "H2", "H3b"):
        sub = full07[full07[f"{sc}_STATUT"].isin(["place", "place_avec_chevauchement"])]
        for r in sub.to_dict("records"):
            rows.append({
                "HYPOTHESE": sc, "ID_SEANCE": r["ID_SEANCE"], "STATUT": r[f"{sc}_STATUT"], "SALLE_ORIGINE": r["NOM_SAL"],
                "CODE_DEST": str(r[f"{sc}_CODE"]).split(".")[0], "LOCAL_DEST": r[f"{sc}_LOCAL"], "DATE": r[f"{sc}_DATE"] or r["DATE"],
                "HDEBUT": r[f"{sc}_DEBUT"] or r["HDEBUT"], "HDEBUT_ORIG": r["HDEBUT"], "HFIN_ORIG": r["HFIN"],
            })
    tr = clean(pd.DataFrame(rows), {"CODE_DEST", "HDEBUT_ORIG", "HFIN_ORIG"})
    tr["CODE_DEST"] = tr["CODE_DEST"].astype(str)
    calc_tr = [
        ("Deb", f"={minutes('HDEBUT')}", ""),
        ("Fin", f"=[@Deb]+({minutes('HFIN_ORIG')})-({minutes('HDEBUT_ORIG')})", ""),
        *hour_formulas(),
        ("H818", "=SUM(t_Rep[@[M08]:[M17]])/60", "num1"),
        ("Sem", "=YEAR([@DATE]+4-WEEKDAY([@DATE],2))*100+ISOWEEKNUM([@DATE])", ""),
        ("Jour", '=INDEX({"lundi","mardi","mercredi","jeudi","vendredi","samedi","dimanche"},WEEKDAY([@DATE],2))', ""),
        ("Idx", "=IFERROR(MATCH([@[CODE_DEST]],Cat_Code,0),0)", ""),
        ("Bat", '=IF([@Idx]=0,"",INDEX(Cat_Bat,[@Idx]))', ""),
        ("Ville", '=IF([@Idx]=0,"",INDEX(Cat_Ville,[@Idx]))', ""),
        ("Site", '=IF([@Idx]=0,"",INDEX(Cat_Site,[@Idx]))', ""),
    ]
    write_table(bk, "Données_Report", "t_Rep", tr, calc_tr, {"LOCAL_DEST": 26, "SALLE_ORIGINE": 26}, tab="#b1781b")
    info["ntr"] = len(tr)

    # --- 28 : tableau de bord par salle et par situation ---
    p28 = src.find("28_tableau_de_bord")
    d28 = clean(read_csv(p28), {"CODE"})
    d28["CODE"] = d28["CODE"].astype(str)
    write_table(bk, "Données_28", "t_28", d28, [("CLE", '=[@SITUATION]&"|"&[@CODE]', "")], {"SALLE": 34, "SITUATION": 24})
    src.register(p28, len(d28), "Listes ; Référence ; H1a ; H1b ; H1c ; H2 (lissage) ; H3b (lissage) ; Amphis")
    info["d28"] = d28

    simple = [
        ("06_surdimensionnement", None, "t_06", "Surdimensionnement", ""),
        ("09_occupation_avant_apres", "Données_09", "t_09", "(aucune)", "Import brut, non utilisé par une feuille de restitution"),
        ("10_classement_occupation", "Données_10", "t_10", "(aucune)", "Import brut, non utilisé par une feuille de restitution"),
        ("11_remplissage_par_salle", "Données_11", "t_11", "(aucune)", "Import brut, non utilisé par une feuille de restitution"),
        ("12_ecarts_effectif_capacite", "Données_12", "t_12", "(aucune)", "Import brut, non utilisé par une feuille de restitution"),
        ("12_ecarts_effectif_capacite_examens", "Données_12X", "t_12X", "(aucune)", "Import brut, non utilisé par une feuille de restitution"),
        ("13_chevauchements_par_salle_semaine", "Données_13", "t_13", "Référence ; H1a ; H1b ; H1c ; H2 (lissage) ; H3b (lissage)", ""),
        ("14_anomalies", "Données_14", "t_14", "(aucune)", "Import brut, non utilisé par une feuille de restitution"),
        ("17_usages_sans_effectif", "Données_17", "t_17", "Usages sans effectif", ""),
        ("18_hypotheses_report", "Données_18", "t_18", "Synthèse V2 ; Conclusion LET ; vues de cas", ""),
        ("23_verification_effectifs_estimes_LET", "Données_23", "t_23", "À vérifier", ""),
        ("24_seances_sans_salle_SSH_Pau", "Données_24", "t_24", "À vérifier", ""),
        ("25_anomalies_capacite", "Données_25", "t_25", "À vérifier", ""),
        ("26_apres_optimisation", "Données_26", "t_26", "Conclusion LET ; Synthèse V2", ""),
        ("27_conclusion_scenarios", "Données_27", "t_27", "Conclusion LET", ""),
    ]
    tables = {}
    for prefix, sheet, table, feed, note in simple:
        path = src.find(prefix)
        df = clean(read_csv(path), {"scenario", "ANOMALIE", "type", "subject", "groups"})
        if table == "t_13":
            df["code"] = df["code"].astype(str)
        if table == "t_24":
            df["exam"] = df["exam"].astype(str)
        tables[table] = df
        if sheet:
            write_table(bk, sheet, table, df, [], {"subject": 30, "groups": 30, "explication": 40, "LIBELLE_MAT": 28, "NOM_DIP": 28, "ANOMALIE": 40, "POURQUOI": 30})
        src.register(path, len(df), feed, note)
    info["tables"] = tables

    # --- 19 à 22 : fichiers par situation, empilés avec une colonne SITUATION ---
    d19 = clean(stacked(src, "19_reaffectation_avant_apres"), {"code", "target_code", "weekday", "type", "subject"})
    d20 = clean(stacked(src, "20_reaffectation_seances"))
    d21 = clean(stacked(src, "21_reaffectation_locaux"), {"code"})
    d22 = clean(stacked(src, "22_effectif_superieur_capacite"))
    if SAMPLE:
        d19, d20 = d19.head(SAMPLE), d20.head(SAMPLE)
    d19["estimated"] = d19["estimated"].astype(str)
    write_table(bk, "Données_19", "t_19", d19, [("CleTri", '=IF([@status]="réaffecté",[@[freed_hours]]+ROW()/10000000,0)', "")], {"room": 30, "subject": 28, "diplomas": 28, "target": 30, "status": 30})
    write_table(bk, "Données_20", "t_20", d20, [], {"LIBELLE_MAT": 28, "NOM_DIP": 28, "LOCAL_INITIAL": 28, "LOCAL_PROPOSE": 28})
    write_table(bk, "Données_21", "t_21", d21, [], {"room": 32})
    write_table(bk, "Données_22", "t_22", d22, [], {"LOCAL": 30, "LIBELLE_MAT": 28, "NOM_DIP": 28})
    feeds = {"19": "Réaffectation globale ; Synthèse V2", "20": "Réaffectation globale (remplissage)", "21": "Réaffectation globale ; Synthèse V2", "22": "Réaffectation globale ; Synthèse V2"}
    for prefix in feeds:
        for name, p in src.files.items():
            if name.startswith(prefix + "_") and name in src.used:
                src.used[name]["sheets"] = feeds[prefix] + " (empilé avec les autres situations)"
    info.update({"d19": d19, "d20": d20, "d21": d21, "d22": d22})
    return info


# =============================================================================================
# 4. Feuilles de service : Listes (listes déroulantes, catalogue des salles) et Calcul (base horaire)
# =============================================================================================
CAT_R0 = 4          # ligne (0-based) de la 1re salle du catalogue dans « Listes »
CAT_C0 = 17         # colonne R : début du catalogue
CALC_R0 = 4         # ligne (0-based) de la 1re salle dans « Calcul »
NROOMS = 596        # remplacé à l'exécution par le nombre réel de salles


def build_listes(bk: Book, info: dict):
    ws = bk.sheets["Listes"]
    d28 = info["d28"]
    ref = d28[d28["SITUATION"] == "Référence"]
    nrooms = len(ref)
    info["nrooms"] = nrooms
    bk.t(ws, 0, 0, "Listes et catalogue des salles", "title")
    bk.t(ws, 1, 0, "Valeurs des listes déroulantes (libellés) et catalogue des salles lu par formule dans le tableau t_28 (situation Référence).", "sub")
    heads = {1: "Périodes", 3: "Situations", 5: "Villes", 7: "Sites", 9: "Bâtiments", 11: "Semaines ISO (clé)", 12: "Semaine"}
    for c, h in heads.items():
        bk.t(ws, 2, c, h, "th")
    for i, p in enumerate(PERIODS):
        ws.write(3 + i, 1, p)
    for i, s in enumerate(SITUATIONS):
        ws.write(3 + i, 3, s)
    for col, key in ((5, "VILLE"), (7, "SITE"), (9, "BATIMENT")):
        values = [ALL] + sorted({str(v) for v in ref[key] if str(v)})
        for i, v in enumerate(values):
            ws.write(3 + i, col, v)
        info[f"n_{key}"] = len(values)
        info[f"list_{key}"] = values
    for i, key in enumerate(info["weeks"]):
        ws.write(3 + i, 11, key)
        ws.write(3 + i, 12, f"{key // 100}-S{key % 100:02d}")
    info["nweeks"] = len(info["weeks"])
    # jours pédagogiques par période, déduits du rapport heures / taux de la situation de référence : J = 10 × Σheures / Σtaux
    bk.t(ws, 2, 14, "Période", "th")
    bk.t(ws, 2, 15, "Jours pédagogiques", "th")
    cols = ["HEURES_ANNEE", "HEURES_S1", "HEURES_S2", "HEURES_EXAMENS", "HEURES_HORS_EXAMENS"]
    rates = ["TAUX_ANNEE_PCT", "TAUX_S1_PCT", "TAUX_S2_PCT", "TAUX_EXAMENS_PCT", "TAUX_HORS_EXAMENS_PCT"]
    for i, p in enumerate(PERIODS):
        ws.write(3 + i, 14, p)
        bk.f(ws, 3 + i, 15,
             f'=IFERROR(ROUND(10*SUMIFS(t_28[{cols[i]}],t_28[SITUATION],"Référence")/SUMIFS(t_28[{rates[i]}],t_28[SITUATION],"Référence"),0),0)')
    ws.write(10, 14, "Début du bloc Référence dans t_28")
    bk.f(ws, 10, 15, '=MATCH("Référence",t_28[SITUATION],0)')
    for c, h in enumerate(["n", "Code", "Salle", "Bât.", "Ville", "Site", "Type", "Capacité", "Amphi LET (1/0)"]):
        bk.t(ws, CAT_R0 - 1, CAT_C0 + c, h, "th")
    start = a1(10, 15)
    for n in range(1, nrooms + 1):
        r = CAT_R0 + n - 1
        ws.write(r, CAT_C0, n)
        idx = f"{start}-1+{a1(r, CAT_C0)}"
        for c, col in enumerate(["CODE", "SALLE", "BATIMENT", "VILLE", "SITE", "TYPE_LOCAL"], start=1):
            bk.f(ws, r, CAT_C0 + c, f'=INDEX(t_28[{col}],{idx})&""', "base")
        bk.f(ws, r, CAT_C0 + 7, f"=N(INDEX(t_28[CAPACITE],{idx}))", "base")
        bk.f(ws, r, CAT_C0 + 8, f'=IF(AND({a1(r, CAT_C0 + 6)}="amphi",{a1(r, CAT_C0 + 3)}="LET"),1,0)', "base")
    last = CAT_R0 + nrooms - 1
    for name, c in (("Cat_Code", 1), ("Cat_Nom", 2), ("Cat_Bat", 3), ("Cat_Ville", 4), ("Cat_Site", 5), ("Cat_Type", 6), ("Cat_Capa", 7), ("Cat_LET", 8)):
        bk.wb.define_name(name, f"=Listes!{rng(CAT_R0, CAT_C0 + c, last, CAT_C0 + c)}")
    bk.wb.define_name("Liste_Periodes", f"=Listes!{rng(3, 1, 7, 1)}")
    bk.wb.define_name("Liste_Situations", f"=Listes!{rng(3, 3, 8, 3)}")
    bk.wb.define_name("Liste_Villes", f"=Listes!{rng(3, 5, 2 + info['n_VILLE'], 5)}")
    bk.wb.define_name("Liste_Sites", f"=Listes!{rng(3, 7, 2 + info['n_SITE'], 7)}")
    bk.wb.define_name("Liste_Batiments", f"=Listes!{rng(3, 9, 2 + info['n_BATIMENT'], 9)}")
    bk.wb.define_name("Jours_Periode", f"=Listes!{rng(3, 15, 7, 15)}")
    bk.wb.define_name("Bloc_Reference", f"=Listes!{a1(10, 15)}")
    ws.set_column(0, 0, 4)
    ws.set_column(1, 15, 16)
    ws.set_column(CAT_C0, CAT_C0 + 8, 18)
    ws.set_column(CAT_C0 + 2, CAT_C0 + 2, 36)
    ws.freeze_panes(3, 0)


def build_calcul(bk: Book, info: dict):
    """Occupation horaire de base de chaque salle (toutes séances du CSV 07) : minutes occupées par tranche de 8 h à 18 h."""
    ws = bk.sheets["Calcul"]
    bk.t(ws, 0, 0, "Calcul : minutes d'occupation par salle et par tranche horaire (situation de référence)", "title")
    bk.t(ws, 1, 0, "SOMME.SI.ENS sur le tableau t_07 (colonnes M08 à M17). Les vues de cas ajoutent ou retirent les séances LET reportées (tableau TR).", "sub")
    bk.t(ws, 3, 0, "n", "th")
    bk.t(ws, 3, 1, "Code", "th")
    for j, h in enumerate(HOURS):
        bk.t(ws, 3, 2 + j, f"{h}h-{h + 1}h", "th")
    for n in range(1, info["nrooms"] + 1):
        r = CALC_R0 + n - 1
        ws.write(r, 0, n)
        bk.f(ws, r, 1, f"=INDEX(Cat_Code,{a1(r, 0)})", "base")
        for j, h in enumerate(HOURS):
            bk.f(ws, r, 2 + j, f"=SUMIFS(t_07[M{h:02d}],t_07[CODE_SAL],{a1(r, 1)})", "num")
    ws.set_column(0, 0, 6)
    ws.set_column(1, 12, 11)
    ws.freeze_panes(4, 2)


# =============================================================================================
# 5. Vues de cas : Référence, H1a, H1b, H1c, H2 (lissage), H3b (lissage)
# =============================================================================================
def bar_chart(bk, ws, kind, title, cats, vals, anchor, size=(560, 330), color=OK, name="", reverse=False, num_fmt=None, legend=False):
    ch = bk.wb.add_chart({"type": kind})
    ch.add_series({"name": name or title, "categories": cats, "values": vals, "fill": {"color": color}, "border": {"none": True},
                   "data_labels": {"value": True, "num_format": num_fmt or "General", "font": {"size": 8}}, "gap": 60})
    ch.set_title({"name": title, "name_font": {"size": 11, "color": OK}})
    ch.set_legend({"none": not legend})
    ch.set_size({"width": size[0], "height": size[1]})
    ch.set_chartarea({"border": {"color": LINE}})
    if reverse:
        ch.set_y_axis({"reverse": True})
    ws.insert_chart(anchor, ch)


def build_view(bk: Book, info: dict, scen: str, sheet: str):
    ws = bk.sheets[sheet]
    n = info["nrooms"]
    is_ref = scen == "Référence"
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 24, 12.5)
    bk.t(ws, 0, 1, f"Occupation des salles · {sheet}", "title")
    if is_ref:
        bk.f(ws, 1, 1, '="Situation actuelle, sans aucun report : occupation issue des séances localisées."', "sub")
    else:
        bk.f(ws, 1, 1, f'=INDEX(t_18[DESCRIPTION],MATCH("{scen}",t_18[HYPOTHESE],0))&" — "&ROUND(INDEX(t_18[SEANCES_A_REPORTER],MATCH("{scen}",t_18[HYPOTHESE],0)),0)&" séances à reporter, "&ROUND(INDEX(t_18[HEURES_A_REPORTER],MATCH("{scen}",t_18[HYPOTHESE],0)),0)&" h"', "sub")
    # --- sélecteurs et cellules de calcul ---
    for c, label in ((1, "Période"), (3, "Ville"), (5, "Site"), (7, "Bâtiment")):
        bk.t(ws, 3, c, label, "label")
    bk.t(ws, 4, 1, "Année", "sel")
    for c in (3, 5, 7):
        bk.t(ws, 4, c, ALL, "sel")
    ws.data_validation(4, 1, 4, 1, {"validate": "list", "source": "=Liste_Periodes"})
    ws.data_validation(4, 3, 4, 3, {"validate": "list", "source": "=Liste_Villes"})
    ws.data_validation(4, 5, 4, 5, {"validate": "list", "source": "=Liste_Sites"})
    ws.data_validation(4, 7, 4, 7, {"validate": "list", "source": "=Liste_Batiments"})
    for c, label in ((9, "n° période"), (10, "critère ville"), (11, "critère site"), (12, "critère bât."), (13, "début bloc t_28")):
        bk.t(ws, 3, c, label, "note")
    P, CV, CS, CB, ST = a1(4, 9), a1(4, 10), a1(4, 11), a1(4, 12), a1(4, 13)
    bk.f(ws, 4, 9, f"=MATCH({a1(4, 1)},Liste_Periodes,0)", "base")
    for c, sel in ((10, a1(4, 3)), (11, a1(4, 5)), (12, a1(4, 7))):
        bk.f(ws, 4, c, f'=IF({sel}="{ALL}","{NOT_MATCH}",{sel})', "base")
    bk.f(ws, 4, 13, f'=MATCH("{scen}",t_28[SITUATION],0)', "base")
    ws.write(3, 14, "")

    # --- tableau de détail par salle (en bas de la feuille) ---
    DH = 175                                  # ligne d'en-tête du tableau de détail
    D0, D1 = DH + 1, DH + n                   # 1re et dernière ligne de données (0-based)
    col = {"salle": 1, "code": 2, "bat": 3, "ville": 4, "site": 5, "type": 6, "capa": 7, "seances": 8, "chev": 9, "heures": 10,
           "taux": 11, "filtre": 12, "opp": 13, "heat": 14}
    R = lambda key: f"${xl_col_to_name(col[key])}${D0 + 1}:${xl_col_to_name(col[key])}${D1 + 1}"
    heads = ["Salle", "Code", "Bât.", "Ville", "Site", "Type", "Capacité", "Séances", "Chevauch.", "Heures", "Taux", "Filtre (1/0)", "Clé optimisation"] + [f"{h}h-{h + 1}h" for h in HOURS]
    bk.t(ws, DH - 1, 1, "Détail par salle et grille horaire (occupation moyenne par tranche d'une heure, année complète)", "h2")
    for c, h in enumerate(heads, start=1):
        bk.t(ws, DH, c, h, "th")
    jours_an = "INDEX(Jours_Periode,1)"
    for i in range(1, n + 1):
        r = DH + i
        idx = f"{ST}-1+{i}"
        bk.f(ws, r, 1, f"=INDEX(t_28[SALLE],{idx})&\"\"", "txt")
        bk.f(ws, r, 2, f"=INDEX(t_28[CODE],{idx})&\"\"", "txt")
        for key, src in (("bat", "BATIMENT"), ("ville", "VILLE"), ("site", "SITE"), ("type", "TYPE_LOCAL")):
            bk.f(ws, r, col[key], f"=INDEX(t_28[{src}],{idx})&\"\"", "txt")
        bk.f(ws, r, col["capa"], f"=N(INDEX(t_28[CAPACITE],{idx}))", "num")
        bk.f(ws, r, col["seances"], f"=INDEX(t_28[SEANCES],{idx})", "num")
        bk.f(ws, r, col["chev"], f"=INDEX(t_28[CHEVAUCHEMENTS],{idx})", "num")
        bk.f(ws, r, col["heures"], f"=INDEX(t_28[[HEURES_ANNEE]:[TAUX_HORS_EXAMENS_PCT]],{idx},2*({P}-1)+1)", "hrs1")
        bk.f(ws, r, col["taux"], f"=INDEX(t_28[[HEURES_ANNEE]:[TAUX_HORS_EXAMENS_PCT]],{idx},2*({P}-1)+2)/100", "pct")
        bat, vil, sit = (a1(r, col[k], False) for k in ("bat", "ville", "site"))
        bk.f(ws, r, col["filtre"],
             f'=IF(AND(OR({a1(4, 3)}="{ALL}",{vil}={a1(4, 3)}),OR({a1(4, 5)}="{ALL}",{sit}={a1(4, 5)}),OR({a1(4, 7)}="{ALL}",{bat}={a1(4, 7)})),1,0)', "num")
        capa, taux, flt = (a1(r, col[k], False) for k in ("capa", "taux", "filtre"))
        bk.f(ws, r, col["opp"], f"=IF(AND({flt}=1,{capa}>=40,{taux}<0.1),{capa}+ROW()/1000000,0)", "base")
        code = a1(r, 2, False)
        for j, h in enumerate(HOURS):
            base = f"Calcul!{a1(CALC_R0 + i - 1, 2 + j)}"
            if is_ref:
                formula = f"={base}/({jours_an}*60)"
            else:
                let = f"INDEX(Cat_LET,{i})"
                formula = (f'=(IF({let}=1,0,{base})+SUMIFS(t_Rep[M{h:02d}],t_Rep[HYPOTHESE],"{scen}",t_Rep[CODE_DEST],{code}))/({jours_an}*60)')
            bk.f(ws, r, col["heat"] + j, formula, "pct")
    ws.conditional_format(D0, col["taux"], D1, col["taux"], {"type": "cell", "criteria": ">", "value": 1, "format": bk.fmt["bad"]})
    ws.conditional_format(D0, col["taux"], D1, col["taux"], {"type": "cell", "criteria": "between", "minimum": 0.8, "maximum": 1, "format": bk.fmt["warn"]})
    ws.conditional_format(D0, col["heat"], D1, col["heat"] + 9, {"type": "2_color_scale", "min_color": "#ffffff", "max_color": OK, "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 1})
    ws.autofilter(DH, 1, D1, col["heat"] + 9)

    # --- KPI ---
    cards = [
        ("Salles suivies", f"=COUNTIFS({R('filtre')},1)", "kpi_n"),
        ("Taux d'occupation moyen", f"=IFERROR(AVERAGEIFS({R('taux')},{R('filtre')},1),0)", "kpi_p"),
        ("Sous-utilisées (< 10 %)", f'=COUNTIFS({R("filtre")},1,{R("taux")},"<"&0.1)', "kpi_n"),
        ("Fortement sollicitées (≥ 30 %)", f'=COUNTIFS({R("filtre")},1,{R("taux")},">="&0.3)', "kpi_n"),
        ("Heures-salles occupées", f"=SUMIFS({R('heures')},{R('filtre')},1)", "kpi_h"),
        ("Séances reconstituées", f"=SUMIFS({R('seances')},{R('filtre')},1)", "kpi_n"),
        ("Chevauchements détectés", ('=SUMIFS(t_28[CHEVAUCHEMENTS],t_28[SITUATION],"Référence")' if is_ref else f'=SUMIFS(t_13[conflicts],t_13[scenario],"{scen}")'), "kpi_n"),
    ]
    for k, (label, formula, fm) in enumerate(cards):
        c = 1 + 2 * k if k else 1
        c0 = 1 if k == 0 else 2 + 2 * (k - 1) + 0
        c0 = 1 + 2 * k - (1 if k else 0) + (0)
        cc = [1, 3, 5, 7, 9, 11, 13][k]
        ws.merge_range(6, cc, 6, cc + 1, label, bk.fmt["kpi_l"])
        ws.merge_range(7, cc, 7, cc + 1, "", bk.fmt["kpi_n"])
        bk.f(ws, 7, cc, formula, fm)
    KS, KT, KH, KSEA, KC = a1(7, 1), a1(7, 3), a1(7, 9), a1(7, 11), a1(7, 13)
    ws.set_row(6, 30)
    ws.set_row(7, 28)
    # synthèse exécutive
    bk.t(ws, 9, 1, "Synthèse exécutive", "h2")
    ex = [("Jours analysés", f"=INDEX(Jours_Periode,{P})", "kpi_n"), ("Créneaux disponibles par salle", f"=INDEX(Jours_Periode,{P})*10", "kpi_h"),
          ("Base quotidienne 8 h-18 h", "=10", "kpi_h"), ("Taux d'occupation moyen", f"={KT}", "kpi_p")]
    for k, (label, formula, fm) in enumerate(ex):
        cc = [1, 3, 5, 7][k]
        ws.merge_range(10, cc, 10, cc + 1, label, bk.fmt["kpi_l"])
        ws.merge_range(11, cc, 11, cc + 1, "", bk.fmt[fm])
        bk.f(ws, 11, cc, formula, fm)
    JOURS = a1(11, 1)

    # --- bâtiments : tableau de travail (colonnes Z:AB) puis classement trié ---
    NB = info["n_BATIMENT"] - 1
    WB0 = 14                                   # 1re ligne du tableau de travail
    wcol = 25
    bk.t(ws, WB0 - 1, wcol, "Bâtiment", "th")
    bk.t(ws, WB0 - 1, wcol + 1, "Taux moyen", "th")
    bk.t(ws, WB0 - 1, wcol + 2, "Clé", "th")
    for i in range(NB):
        r = WB0 + i
        bk.f(ws, r, wcol, f"=INDEX(Liste_Batiments,{i + 2})", "base")
        bk.f(ws, r, wcol + 1, f"=IFERROR(AVERAGEIFS({R('taux')},{R('bat')},{a1(r, wcol)},{R('filtre')},1),-1)", "pct")
        bk.f(ws, r, wcol + 2, f"={a1(r, wcol + 1)}+ROW()/1000000000", "base")
    WN, WV, WK = (rng(WB0, wcol + k, WB0 + NB - 1, wcol + k) for k in range(3))
    TOP = 20
    bk.t(ws, 13, 1, f"Taux moyen par bâtiment (classement des {TOP} premiers)", "h2")
    bk.t(ws, 14, 1, "Bâtiment", "th")
    bk.t(ws, 14, 2, "Taux moyen", "th")
    for k in range(1, TOP + 1):
        r = 14 + k
        key = f"LARGE({WK},{k})"
        bk.f(ws, r, 1, f'=IF({key}<0,"",INDEX({WN},MATCH({key},{WK},0)))', "txt")
        bk.f(ws, r, 2, f'=IF({key}<0,0,INDEX({rng(WB0, wcol + 1, WB0 + NB - 1, wcol + 1)},MATCH({key},{WK},0)))', "pct")
    bar_chart(bk, ws, "bar", "Taux moyen par bâtiment", f"='{sheet}'!{rng(15, 1, 14 + TOP, 1)}", f"='{sheet}'!{rng(15, 2, 14 + TOP, 2)}", "E14", (620, 380), reverse=True, num_fmt="0.0%")

    # --- répartition par tranche ---
    r0 = 37
    bk.t(ws, r0, 1, "Répartition des salles par tranche de taux d'occupation", "h2")
    for c, h in enumerate(["Tranche", "Salles", "Part"]):
        bk.t(ws, r0 + 1, 1 + c, h, "th")
    labels = [f"{a}–{b} %" if b < 10**5 else "40 % et plus" for a, b in BUCKETS]
    for k, ((lo, hi), lab) in enumerate(zip(BUCKETS, labels)):
        r = r0 + 2 + k
        bk.t(ws, r, 1, lab, "txt")
        bk.f(ws, r, 2, f'=COUNTIFS({R("filtre")},1,{R("taux")},">="&{lo / 100},{R("taux")},"<"&{hi / 100})', "num")
        bk.f(ws, r, 3, f"=IF({KS}=0,0,{a1(r, 2)}/{KS})", "pct")
    bar_chart(bk, ws, "column", "Nombre de salles par tranche de taux", f"='{sheet}'!{rng(r0 + 2, 1, r0 + 9, 1)}", f"='{sheet}'!{rng(r0 + 2, 2, r0 + 9, 2)}", "F38", (520, 260))

    # --- salles à optimiser : ≥ 40 places occupées < 10 % ---
    r1 = r0 + 12
    bk.t(ws, r1, 1, "Fort potentiel d'optimisation : salles de 40 places ou plus occupées moins de 10 % du temps", "h2")
    for c, h in enumerate(["Salle", "Places", "Taux"]):
        bk.t(ws, r1 + 1, 1 + c, h, "th")
    for k in range(1, 9):
        r = r1 + 1 + k
        key = f"LARGE({R('opp')},{k})"
        m = f"MATCH({key},{R('opp')},0)"
        bk.f(ws, r, 1, f'=IF({key}=0,"",INDEX({R("salle")},{m}))', "txt")
        bk.f(ws, r, 2, f'=IF({key}=0,"",INDEX({R("capa")},{m}))', "num")
        bk.f(ws, r, 3, f'=IF({key}=0,"",INDEX({R("taux")},{m}))', "pct")

    # --- occupation hebdomadaire ---
    r2 = r1 + 12
    NW = info["nweeks"]
    bk.t(ws, r2, 1, "Occupation moyenne par semaine ISO (année complète)", "h2")
    for c, h in enumerate(["Semaine", "Clé", "Heures", "Jours", "Taux", "Clé tri"]):
        bk.t(ws, r2 + 1, 1 + c, h, "th")
    flt_t07 = f"t_07[Ville],{CV},t_07[Site],{CS},t_07[Bat],{CB}"
    flt_tr = f"t_Rep[Ville],{CV},t_Rep[Site],{CS},t_Rep[Bat],{CB}"
    for k in range(NW):
        r = r2 + 2 + k
        bk.f(ws, r, 1, f"=INDEX(Listes!{rng(3, 12, 2 + NW, 12)},{k + 1})", "txt")
        bk.f(ws, r, 2, f"=INDEX(Listes!{rng(3, 11, 2 + NW, 11)},{k + 1})", "base")
        key = a1(r, 2)
        base = f"SUMIFS(t_07[H818],t_07[Sem],{key},{flt_t07})"
        if is_ref:
            hours = f"={base}"
        else:
            hours = f'={base}-SUMIFS(t_07[H818],t_07[Sem],{key},t_07[SrcLET],1,{flt_t07})+SUMIFS(t_Rep[H818],t_Rep[HYPOTHESE],"{scen}",t_Rep[Sem],{key},{flt_tr})'
        bk.f(ws, r, 3, hours, "hrs1")
        bk.f(ws, r, 4, f"=SUMIFS(t_07[JourFlag],t_07[Sem],{key})", "num")
        bk.f(ws, r, 5, f"=IF({KS}*{a1(r, 4)}=0,0,{a1(r, 3)}/({KS}*{a1(r, 4)}*10))", "pct")
        bk.f(ws, r, 6, f"={a1(r, 5)}+ROW()/1000000000", "base")
    wk0, wk1 = r2 + 2, r2 + 1 + NW
    bar_chart(bk, ws, "column", "Taux d'occupation par semaine", f"='{sheet}'!{rng(wk0, 1, wk1, 1)}", f"='{sheet}'!{rng(wk0, 5, wk1, 5)}", "I" + str(r2 + 2), (720, 300), num_fmt="0%")
    # semaines de pointe
    rp = wk1 + 2
    bk.t(ws, rp, 1, "5 semaines les plus chargées", "h2")
    for c, h in enumerate(["Semaine", "Taux"]):
        bk.t(ws, rp + 1, 1 + c, h, "th")
    for k in range(1, 6):
        r = rp + 1 + k
        key = f"LARGE({rng(wk0, 6, wk1, 6)},{k})"
        bk.f(ws, r, 1, f'=IFERROR(INDEX({rng(wk0, 1, wk1, 1)},MATCH({key},{rng(wk0, 6, wk1, 6)},0)),"")', "txt")
        bk.f(ws, r, 2, f'=IFERROR(INDEX({rng(wk0, 5, wk1, 5)},MATCH({key},{rng(wk0, 6, wk1, 6)},0)),"")', "pct")

    # --- heatmap jour × heure ---
    rh = rp + 9
    bk.t(ws, rh, 1, "Occupation moyenne par jour de la semaine et par tranche horaire (filtres actifs, année complète)", "h2")
    bk.t(ws, rh + 1, 1, "Jour", "th")
    for j, h in enumerate(HOURS):
        bk.t(ws, rh + 1, 2 + j, f"{h}h-{h + 1}h", "th")
    bk.t(ws, rh + 1, 12, "Moyenne", "th")
    for d, day in enumerate(DAYS):
        r = rh + 2 + d
        bk.t(ws, r, 1, day.capitalize(), "txt_b")
        for j, h in enumerate(HOURS):
            base = f'SUMIFS(t_07[M{h:02d}],t_07[JOUR],"{day}",{flt_t07})'
            if is_ref:
                total = base
            else:
                total = f'{base}-SUMIFS(t_07[M{h:02d}],t_07[JOUR],"{day}",t_07[SrcLET],1,{flt_t07})+SUMIFS(t_Rep[M{h:02d}],t_Rep[HYPOTHESE],"{scen}",t_Rep[Jour],"{day}",{flt_tr})'
            ndays = f'SUMIFS(t_07[JourFlag],t_07[JOUR],"{day}")'
            bk.f(ws, r, 2 + j, f"=IF({KS}*{ndays}=0,0,({total})/({KS}*{ndays}*60))", "pct")
        bk.f(ws, r, 12, f"=AVERAGE({rng(r, 2, r, 11)})", "pct")
    ws.conditional_format(rh + 2, 2, rh + 6, 11, {"type": "2_color_scale", "min_color": "#ffffff", "max_color": OK, "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 1})
    bar_chart(bk, ws, "column", "Jour de pointe : taux moyen par jour", f"='{sheet}'!{rng(rh + 2, 1, rh + 6, 1)}", f"='{sheet}'!{rng(rh + 2, 12, rh + 6, 12)}", "O" + str(rh), (420, 250), color=WARN, num_fmt="0.0%")

    # --- top / flop bâtiments ---
    rt = rh + 9
    bk.t(ws, rt, 1, "Top 5 bâtiments", "h2")
    bk.t(ws, rt, 5, "5 bâtiments les plus faibles", "h2")
    for c, h in enumerate(["Bâtiment", "Taux"]):
        bk.t(ws, rt + 1, 1 + c, h, "th")
        bk.t(ws, rt + 1, 5 + c, h, "th")
    for k in range(1, 6):
        r = rt + 1 + k
        top = f"LARGE({WK},{k})"
        bk.f(ws, r, 1, f'=IF({top}<0,"",INDEX({WN},MATCH({top},{WK},0)))', "txt")
        bk.f(ws, r, 2, f'=IF({top}<0,"",INDEX({rng(WB0, wcol + 1, WB0 + NB - 1, wcol + 1)},MATCH({top},{WK},0)))', "pct")
        low = f"AGGREGATE(15,6,{WK}/({rng(WB0, wcol + 1, WB0 + NB - 1, wcol + 1)}>=0),{k})"
        bk.f(ws, r, 5, f'=IFERROR(INDEX({WN},MATCH({low},{WK},0)),"")', "txt")
        bk.f(ws, r, 6, f'=IFERROR(INDEX({rng(WB0, wcol + 1, WB0 + NB - 1, wcol + 1)},MATCH({low},{WK},0)),"")', "pct")

    # --- chevauchements ---
    rc = rt + 9
    bk.t(ws, rc, 1, "Chevauchements : par salle et par semaine (40 premières lignes)", "h2")
    if is_ref:
        bk.t(ws, rc + 1, 1, "Situation de référence : les 179 doubles réservations existantes ne sont pas détaillées par semaine dans les CSV du dossier (voir CSV 07).", "note")
    else:
        for c, h in enumerate(["Salle", "Semaine", "Période", "Chevauchements"]):
            bk.t(ws, rc + 1, 1 + c, h, "th")
        for k in range(1, 41):
            r = rc + 1 + k
            row = f'AGGREGATE(15,6,(ROW(t_13[scenario])-ROW(t_13[[#Headers],[scenario]]))/(t_13[scenario]="{scen}"),{k})'
            for c, colname in enumerate(["room", "week", "period", "conflicts"]):
                bk.f(ws, r, 1 + c, f'=IFERROR(INDEX(t_13[{colname}],{row}),"")', "txt" if c < 3 else "num")
        ws.conditional_format(rc + 2, 4, rc + 41, 4, {"type": "cell", "criteria": ">", "value": 0, "format": bk.fmt["bad"]})
    assert rc + 43 < DH - 2, (rc, DH)
    ws.freeze_panes(6, 0)
    ws.set_zoom(90)
    return {"kpi_row": 7, "cols": cc}


# =============================================================================================
# 6. Onglets : Amphis, Réaffectation globale, Synthèse V2, Conclusion LET, À vérifier, Usages, Surdimensionnement
# =============================================================================================
CASES = ["Référence", "H1a", "H1b", "H1c", "H2", "H3b", "Réaffectation (référence)"]
CASE_LABELS = ["Référence", "H1a", "H1b", "H1c", "H2", "H3b", "Réaffectation"]


def cell_pair(bk, ws, r, c, label_formula, value_formula, fmt):
    bk.f(ws, r, c, label_formula, "txt")
    bk.f(ws, r, c + 1, value_formula, fmt)


def header_row(bk, ws, r, c0, labels, height=30):
    for i, h in enumerate(labels):
        bk.t(ws, r, c0 + i, h, "th")
    ws.set_row(r, height)


def build_amphis(bk: Book, info: dict):
    ws = bk.sheets["Amphis"]
    d28 = info["d28"]
    am = d28[(d28["SITUATION"] == "Référence") & (d28["TYPE_LOCAL"] == "amphi")].sort_values(["BATIMENT", "SALLE"])
    na = len(am)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 10)
    ws.set_column(2, 2, 26)
    ws.set_column(3, 40, 11)
    bk.t(ws, 0, 1, "Amphithéâtres", "title")
    bk.t(ws, 1, 1, "Taux d'occupation (8 h-18 h) de chaque amphi avant (Référence) et après report. Cochez des amphis (x dans la colonne Sélection) pour calculer leur moyenne.", "sub")
    for c, label in ((1, "Période"), (3, "Ville"), (5, "Site"), (7, "Bâtiment")):
        bk.t(ws, 3, c, label, "label")
    bk.t(ws, 4, 1, "Année", "sel")
    for c in (3, 5, 7):
        bk.t(ws, 4, c, ALL, "sel")
    ws.data_validation(4, 1, 4, 1, {"validate": "list", "source": ["Année", "Examens", "Hors examens"]})
    ws.data_validation(4, 3, 4, 3, {"validate": "list", "source": "=Liste_Villes"})
    ws.data_validation(4, 5, 4, 5, {"validate": "list", "source": "=Liste_Sites"})
    ws.data_validation(4, 7, 4, 7, {"validate": "list", "source": "=Liste_Batiments"})
    bk.t(ws, 3, 9, "n° période t_28", "note")
    P = a1(4, 9)
    bk.f(ws, 4, 9, f'=CHOOSE(MATCH({a1(4, 1)},{{"Année","Examens","Hors examens"}},0),1,4,5)', "base")
    TH = 14
    T0, T1 = TH + 1, TH + na
    c_sel, c_name, c_code, c_bat, c_vil, c_sit, c_cap, c_flt, c_c0 = 1, 2, 3, 4, 5, 6, 7, 8, 9
    R = lambda c: f"${xl_col_to_name(c)}${T0 + 1}:${xl_col_to_name(c)}${T1 + 1}"
    bk.t(ws, TH - 1, 1, "Détail par amphi", "h2")
    labels = ["Sélection", "Amphi", "Code", "Bât.", "Ville", "Site", "Places", "Filtre"]
    for i, lab in enumerate(CASE_LABELS):
        labels += [f"Taux {lab}", f"Heures {lab}"]
    header_row(bk, ws, TH, 1, labels)
    ws.data_validation(T0, c_sel, T1, c_sel, {"validate": "list", "source": ["x"], "ignore_blank": True})
    for i, (_, room) in enumerate(am.iterrows()):
        r = T0 + i
        ws.write(r, c_code, str(room["CODE"]))
        ws.write(r, c_sel, "", bk.fmt["sel"])
        bk.f(ws, r, c_name, f'=IFERROR(INDEX(t_28[SALLE],MATCH("Référence|"&{a1(r, c_code, False)},t_28[CLE],0))&"","")', "txt")
        for c, key in ((c_bat, "BATIMENT"), (c_vil, "VILLE"), (c_sit, "SITE")):
            bk.f(ws, r, c, f'=INDEX(t_28[{key}],MATCH("Référence|"&{a1(r, c_code, False)},t_28[CLE],0))&""', "txt")
        bk.f(ws, r, c_cap, f'=N(INDEX(t_28[CAPACITE],MATCH("Référence|"&{a1(r, c_code, False)},t_28[CLE],0)))', "num")
        bk.f(ws, r, c_flt,
             f'=IF(AND(OR({a1(4, 3)}="{ALL}",{a1(r, c_vil, False)}={a1(4, 3)}),OR({a1(4, 5)}="{ALL}",{a1(r, c_sit, False)}={a1(4, 5)}),OR({a1(4, 7)}="{ALL}",{a1(r, c_bat, False)}={a1(4, 7)})),1,0)', "num")
        for k, case in enumerate(CASES):
            m = f'MATCH("{case}|"&{a1(r, c_code, False)},t_28[CLE],0)'
            bk.f(ws, r, c_c0 + 2 * k, f"=INDEX(t_28[[HEURES_ANNEE]:[TAUX_HORS_EXAMENS_PCT]],{m},2*({P}-1)+2)/100", "pct")
            bk.f(ws, r, c_c0 + 2 * k + 1, f"=INDEX(t_28[[HEURES_ANNEE]:[TAUX_HORS_EXAMENS_PCT]],{m},2*({P}-1)+1)", "hrs1")
    ws.autofilter(TH, 1, T1, c_c0 + 13)
    ws.freeze_panes(TH + 1, 3)
    # bloc moyenne générale
    bk.t(ws, 6, 1, "Moyenne générale des amphis sélectionnés (ou affichés si aucun n'est coché)", "h2")
    bk.t(ws, 7, 1, "Amphis retenus", "txt_b")
    bk.f(ws, 7, 2, f'=IF(COUNTIF({R(c_sel)},"x")>0,COUNTIF({R(c_sel)},"x"),COUNTIFS({R(c_flt)},1))', "kpi_n")
    header_row(bk, ws, 8, 1, ["Indicateur"] + CASE_LABELS, 18)
    rows = ["Taux moyen", "Écart vs Référence (pts)", "Heures totales", "Heures moyennes par amphi"]
    for i, lab in enumerate(rows):
        bk.t(ws, 9 + i, 1, lab, "txt_b")
    for k in range(7):
        tc, hc = c_c0 + 2 * k, c_c0 + 2 * k + 1
        cond = f'IF(COUNTIF({R(c_sel)},"x")>0,{{}},{{}})'
        avg = f'=IFERROR(IF(COUNTIF({R(c_sel)},"x")>0,AVERAGEIFS({R(tc)},{R(c_sel)},"x"),AVERAGEIFS({R(tc)},{R(c_flt)},1)),0)'
        tot = f'=IF(COUNTIF({R(c_sel)},"x")>0,SUMIFS({R(hc)},{R(c_sel)},"x"),SUMIFS({R(hc)},{R(c_flt)},1))'
        bk.f(ws, 9, 2 + k, avg, "pct")
        bk.f(ws, 10, 2 + k, f"=({a1(9, 2 + k)}-{a1(9, 2)})*100", "num1")
        bk.f(ws, 11, 2 + k, tot, "hrs")
        bk.f(ws, 12, 2 + k, f"=IF({a1(7, 2)}=0,0,{a1(11, 2 + k)}/{a1(7, 2)})", "hrs1")
    bar_chart(bk, ws, "column", "Taux moyen des amphis retenus, par cas", f"='Amphis'!{rng(8, 2, 8, 8)}", f"='Amphis'!{rng(9, 2, 9, 8)}", "L5", (560, 250), num_fmt="0.0%")
    ws.conditional_format(T0, c_c0, T1, c_c0 + 13, {"type": "cell", "criteria": ">", "value": 1, "format": bk.fmt["bad"]})
    ws.set_zoom(90)


def realloc_kpis(sit: str, ref: str) -> dict[str, str]:
    """Formules des indicateurs de réaffectation pour une situation (cellule `ref` = nom de la situation)."""
    s19 = f"t_19[SITUATION],{ref}"
    out = {
        "analysees": f"=COUNTIFS(t_20[SITUATION],{ref})",
        "moved_n": f'=SUMIFS(t_19[sessions],{s19},t_19[status],"réaffecté")',
        "moved_h": f'=SUMIFS(t_19[freed_hours],{s19},t_19[status],"réaffecté")',
        "out_h": f'=SUMIFS(t_19[freed_hours],{s19},t_19[status],"réaffecté",t_19[kind],"amphi")',
        "in_h": f'=SUMIFS(t_19[freed_hours],{s19},t_19[status],"réaffecté",t_19[target_kind],"amphi")',
        "net_h": "",
        "occ_b": f'=IFERROR(AVERAGEIFS(t_21[rate_before],t_21[SITUATION],{ref},t_21[kind],"amphi")/100,0)',
        "occ_a": f'=IFERROR(AVERAGEIFS(t_21[rate_after],t_21[SITUATION],{ref},t_21[kind],"amphi")/100,0)',
        "fill_b": f'=IFERROR(SUMPRODUCT((t_20[SITUATION]={ref})*(t_20[EFFECTIF]<>"")*t_20[EFFECTIF]*t_20[HEURES])/SUMPRODUCT((t_20[SITUATION]={ref})*(t_20[EFFECTIF]<>"")*t_20[CAPACITE_INITIALE]*t_20[HEURES]),0)',
        "fill_a": f'=IFERROR(SUMPRODUCT((t_20[SITUATION]={ref})*(t_20[EFFECTIF]<>"")*t_20[EFFECTIF]*t_20[HEURES])/SUMPRODUCT((t_20[SITUATION]={ref})*(t_20[EFFECTIF]<>"")*t_20[CAPACITE_PROPOSEE]*t_20[HEURES]),0)',
        "over": f"=COUNTIFS(t_22[SITUATION],{ref})",
        "est_n": f'=SUMIFS(t_19[sessions],{s19},t_19[status],"réaffecté",t_19[estimated],"True")',
        "est_h": f'=SUMIFS(t_19[freed_hours],{s19},t_19[status],"réaffecté",t_19[estimated],"True")',
        "unk_n": f'=SUMIFS(t_19[sessions],{s19},t_19[status],"effectif inconnu : non réaffectable")',
        "unk_h": f'=SUMIFS(t_19[hours],{s19},t_19[status],"effectif inconnu : non réaffectable")',
    }
    return out


def build_realloc(bk: Book, info: dict):
    ws = bk.sheets["Réaffectation globale"]
    d19 = info["d19"]
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 34)
    ws.set_column(2, 16, 13)
    bk.t(ws, 0, 1, "Réaffectation selon les effectifs", "title")
    bk.t(ws, 1, 1, "Cours et TD du collège SSH à Pau : on compare l'effectif à la taille de la salle et on propose une salle plus adaptée, au même jour et à la même heure, sans créer de conflit.", "sub")
    bk.t(ws, 3, 1, "Situation", "label")
    bk.t(ws, 4, 1, "Référence", "sel")
    ws.data_validation(4, 1, 4, 1, {"validate": "list", "source": "=Liste_Situations"})
    SEL = a1(4, 1)
    k = realloc_kpis("", SEL)
    cards = [
        ("Cours/TD SSH Pau analysés", k["analysees"], "kpi_n"), ("Séances réaffectées", k["moved_n"], "kpi_n"), ("Heures réaffectées (8 h-18 h)", k["moved_h"], "kpi_h"),
        ("Heures quittant des amphis", k["out_h"], "kpi_h"), ("Heures d'amphi libérées (net)", f"={a1(7, 7)}-({k['in_h'][1:]})", "kpi_h"),
        ("Occupation amphis : avant", k["occ_b"], "kpi_p"), ("Occupation amphis : après", k["occ_a"], "kpi_p"),
        ("Remplissage : avant", k["fill_b"], "kpi_p"), ("Remplissage : après", k["fill_a"], "kpi_p"),
        ("Effectif > capacité (séances)", k["over"], "kpi_n"), ("dont fondées sur un effectif estimé (séances)", k["est_n"], "kpi_n"),
        ("Sans effectif, non réaffectables (séances)", k["unk_n"], "kpi_n"), ("Sans effectif (heures)", k["unk_h"], "kpi_h"),
    ]
    for i, (label, formula, fm) in enumerate(cards):
        row = 6 if i < 7 else 9
        c = 1 + (i if i < 7 else i - 7) * 2 if False else None
    positions = [(6, 1), (6, 3), (6, 5), (6, 7), (6, 9), (6, 11), (6, 13), (9, 1), (9, 3), (9, 5), (9, 7), (9, 9), (9, 11)]
    for (label, formula, fm), (r, c) in zip(cards, positions):
        ws.merge_range(r, c, r, c + 1, label, bk.fmt["kpi_l"])
        ws.merge_range(r + 1, c, r + 1, c + 1, "", bk.fmt[fm])
        bk.f(ws, r + 1, c, formula, fm)
    ws.set_row(6, 30)
    ws.set_row(9, 30)
    # heures d'amphi libérées (net) : heures quittant - heures entrant, 2e carte de la ligne 7
    bk.f(ws, 7, 7, k["out_h"], "kpi_h")
    bk.f(ws, 7, 9, f"={a1(7, 7)}-({k['in_h'][1:]})", "kpi_h")
    # --- tableau avant / après (100 premières séries réaffectées, par heures libérées décroissantes) ---
    r0 = 13
    NTOP = 100
    bk.t(ws, r0, 1, "Tableau avant / après : séries réaffectées, classées par heures libérées", "h2")
    header_row(bk, ws, r0 + 1, 1, ["Local actuel", "Bât.", "Capacité", "Effectif max", "Local proposé", "Capacité proposée", "Créneau", "Matière", "Promotions", "Séances", "Heures libérées"])
    for i in range(1, NTOP + 1):
        r = r0 + 1 + i
        key = f"AGGREGATE(14,6,t_19[CleTri]/(t_19[SITUATION]={SEL}),{i})"
        m = f"MATCH({key},t_19[CleTri],0)"
        g = lambda col: f"INDEX(t_19[{col}],{m})"
        bk.f(ws, r, 1, f'=IFERROR(IF({key}<=0,"",{g("room")}),"")', "txt")
        bk.f(ws, r, 2, f'=IFERROR(IF({key}<=0,"",{g("building")}),"")', "txt")
        bk.f(ws, r, 3, f'=IFERROR(IF({key}<=0,"",{g("capacity")}),"")', "num")
        bk.f(ws, r, 4, f'=IFERROR(IF({key}<=0,"",{g("max_effect")}&IF({g("estimated")}="True"," (estimé)","")),"")', "txt")
        bk.f(ws, r, 5, f'=IFERROR(IF({key}<=0,"",{g("target")}),"")', "txt")
        bk.f(ws, r, 6, f'=IFERROR(IF({key}<=0,"",{g("target_capacity")}),"")', "num")
        bk.f(ws, r, 7, f'=IFERROR(IF({key}<=0,"",{g("weekday")}&" "&INT({g("start")}/60)&"h"&TEXT(MOD({g("start")},60),"00")&"-"&INT({g("end")}/60)&"h"&TEXT(MOD({g("end")},60),"00")),"")', "txt")
        bk.f(ws, r, 8, f'=IFERROR(IF({key}<=0,"",{g("subject")}&" ("&{g("type")}&")"),"")', "txt")
        bk.f(ws, r, 9, f'=IFERROR(IF({key}<=0,"",{g("diplomas")}),"")', "txt")
        bk.f(ws, r, 10, f'=IFERROR(IF({key}<=0,"",{g("sessions")}),"")', "num")
        bk.f(ws, r, 11, f'=IFERROR(IF({key}<=0,"",{g("freed_hours")}),"")', "hrs1")
    ws.set_column(8, 9, 30)
    # --- séries non réaffectées ---
    r1 = r0 + NTOP + 4
    bk.t(ws, r1, 1, "Séries non réaffectées", "h2")
    header_row(bk, ws, r1 + 1, 1, ["Raison", "Séries", "Séances", "Heures (8 h-18 h)"])
    statuses = sorted(s for s in d19["status"].unique() if s != "réaffecté")
    for i, st in enumerate(statuses):
        r = r1 + 2 + i
        bk.t(ws, r, 1, st, "txt")
        bk.f(ws, r, 2, f'=COUNTIFS(t_19[SITUATION],{SEL},t_19[status],{a1(r, 1)})', "num")
        bk.f(ws, r, 3, f'=SUMIFS(t_19[sessions],t_19[SITUATION],{SEL},t_19[status],{a1(r, 1)})', "num")
        bk.f(ws, r, 4, f'=SUMIFS(t_19[hours],t_19[SITUATION],{SEL},t_19[status],{a1(r, 1)})', "hrs")
    # --- séances dont l'effectif dépasse la capacité (50 premières) ---
    r2 = r1 + len(statuses) + 4
    bk.t(ws, r2, 1, "Séances dont l'effectif dépasse la capacité (50 premières ; liste complète dans Données_22)", "h2")
    header_row(bk, ws, r2 + 1, 1, ["Date", "Local", "Capacité", "Effectif", "Dépassement", "Type", "Matière", "Examen"])
    for i in range(1, 51):
        r = r2 + 1 + i
        row = f"AGGREGATE(15,6,(ROW(t_22[SITUATION])-ROW(t_22[[#Headers],[SITUATION]]))/(t_22[SITUATION]={SEL}),{i})"
        for c, (col, fm) in enumerate([("DATE", "date"), ("LOCAL", "txt"), ("CAPACITE", "num"), ("EFFECTIF", "num"), ("DEPASSEMENT", "num"), ("TYPE", "txt"), ("LIBELLE_MAT", "txt"), ("EXAMEN", "txt")]):
            bk.f(ws, r, 1 + c, f'=IFERROR(INDEX(t_22[{col}],{row}),"")', fm)
    # --- locaux avant / après (amphis) ---
    r3 = r2 + 55
    bk.t(ws, r3, 1, "Amphis de Pau : taux d'occupation avant / après (graphique)", "h2")
    header_row(bk, ws, r3 + 1, 1, ["Amphi", "Taux avant", "Taux après", "Variation (pts)", "Remplissage avant", "Remplissage après"])
    namph = int((info["d21"][(info["d21"]["SITUATION"] == "Référence") & (info["d21"]["kind"] == "amphi")]).shape[0])
    for i in range(1, namph + 1):
        r = r3 + 1 + i
        row = f'AGGREGATE(15,6,(ROW(t_21[SITUATION])-ROW(t_21[[#Headers],[SITUATION]]))/((t_21[SITUATION]={SEL})*(t_21[kind]="amphi")),{i})'
        bk.f(ws, r, 1, f'=IFERROR(INDEX(t_21[room],{row}),"")', "txt")
        bk.f(ws, r, 2, f'=IFERROR(INDEX(t_21[rate_before],{row})/100,0)', "pct")
        bk.f(ws, r, 3, f'=IFERROR(INDEX(t_21[rate_after],{row})/100,0)', "pct")
        bk.f(ws, r, 4, f"=({a1(r, 3)}-{a1(r, 2)})*100", "num1")
        bk.f(ws, r, 5, f'=IFERROR(N(INDEX(t_21[fill_before],{row}))/100,0)', "pct")
        bk.f(ws, r, 6, f'=IFERROR(N(INDEX(t_21[fill_after],{row}))/100,0)', "pct")
    ch = bk.wb.add_chart({"type": "column"})
    for c, (nm, color) in enumerate([("Avant", "#9aa7b1"), ("Après", OK)], start=2):
        ch.add_series({"name": nm, "categories": f"='Réaffectation globale'!{rng(r3 + 2, 1, r3 + 1 + namph, 1)}", "values": f"='Réaffectation globale'!{rng(r3 + 2, c, r3 + 1 + namph, c)}", "fill": {"color": color}, "gap": 80})
    ch.set_title({"name": "Taux d'occupation des amphis avant / après réaffectation", "name_font": {"size": 11, "color": OK}})
    ch.set_size({"width": 820, "height": 330})
    ch.set_y_axis({"num_format": "0%"})
    ch.set_x_axis({"num_font": {"rotation": -45, "size": 8}})
    ws.insert_chart(r3 + 1, 8, ch)
    ws.conditional_format(r3 + 2, 4, r3 + 1 + namph, 4, {"type": "cell", "criteria": "<", "value": 0, "format": bk.fmt["good"]})
    ws.conditional_format(r2 + 2, 5, r2 + 51, 5, {"type": "cell", "criteria": ">", "value": 0, "format": bk.fmt["bad"]})
    ws.freeze_panes(6, 0)
    ws.set_zoom(90)


def build_synthese(bk: Book, info: dict):
    ws = bk.sheets["Synthèse V2"]
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 46)
    ws.set_column(2, 14, 15)
    bk.t(ws, 0, 1, "Synthèse V2", "title")
    bk.t(ws, 1, 1, "Hypothèses de report des trois amphis de Lettres (H1a, H1b, H1c) et comparaisons avec lissage (H2, H3b). Heures comptées entre 8 h et 18 h.", "sub")
    H = lambda col, key: f'INDEX(t_18[{col}],MATCH("{key}",t_18[HYPOTHESE],0))'
    bk.t(ws, 3, 1, "L'essentiel", "h2")
    lines = [
        f'="Les trois amphis de Lettres représentent "&ROUND({H("SEANCES_A_REPORTER", "H1a")},0)&" séances ("&ROUND({H("HEURES_A_REPORTER", "H1a")},0)&" h) à reporter, dont "&ROUND({H("DONT_EXAMENS_SEANCES", "H1a")},0)&" examens ou contrôles ("&ROUND({H("DONT_EXAMENS_HEURES", "H1a")},0)&" h)."',
    ]
    for key in ("H1a", "H1b", "H1c"):
        lines.append(
            f'="{key} : "&ROUND({H("REPORTEES_EFFECTIF_CONNU_SEANCES", key)}+{H("REPORTEES_EFFECTIF_ESTIME_SEANCES", key)},0)&" séances reportées, dont "&ROUND({H("DONT_AVEC_CHEVAUCHEMENT_SEANCES", key)},0)&" avec chevauchement ("&ROUND({H("CHEVAUCHEMENTS", key)},0)&" moments en double). Besoin non résolu : "&ROUND({H("BESOIN_NON_RESOLU_SEANCES", key)},0)&" séances, "&ROUND({H("BESOIN_NON_RESOLU_HEURES", key)},0)&" h."'
        )
    lines.append(
        '="Réaffectation globale (référence) : "&ROUND(SUMIFS(t_19[sessions],t_19[SITUATION],"Référence",t_19[status],"réaffecté"),0)&" séances de cours/TD ("&ROUND(SUMIFS(t_19[freed_hours],t_19[SITUATION],"Référence",t_19[status],"réaffecté"),0)&" h) peuvent aller dans un local plus adapté, sans aucun conflit."'
    )
    for i, line in enumerate(lines):
        ws.merge_range(4 + i, 1, 4 + i, 12, "", bk.fmt["txt"])
        bk.f(ws, 4 + i, 1, line, "txt")
        ws.set_row(4 + i, 30)
    r0 = 4 + len(lines) + 1
    bk.t(ws, r0, 1, "Hypothèses de report des 3 amphis de Lettres", "h2")
    cols = [("Locaux d'accueil", "LOCAUX_ACCUEIL", None), ("À reporter", "SEANCES_A_REPORTER", "HEURES_A_REPORTER"), ("dont examens", "DONT_EXAMENS_SEANCES", "DONT_EXAMENS_HEURES"),
            ("Reportées (effectif connu)", "REPORTEES_EFFECTIF_CONNU_SEANCES", "REPORTEES_EFFECTIF_CONNU_HEURES"), ("Reportées (effectif estimé)", "REPORTEES_EFFECTIF_ESTIME_SEANCES", "REPORTEES_EFFECTIF_ESTIME_HEURES"),
            ("dont avec chevauchement", "DONT_AVEC_CHEVAUCHEMENT_SEANCES", "DONT_AVEC_CHEVAUCHEMENT_HEURES"), ("Chevauchements", "CHEVAUCHEMENTS", None),
            ("Sans effectif", "SANS_EFFECTIF_SEANCES", "SANS_EFFECTIF_HEURES"), ("Sans solution", "SANS_SOLUTION_SEANCES", "SANS_SOLUTION_HEURES"),
            ("Non couvertes (capacité)", "NON_COUVERTES_SEANCES", "NON_COUVERTES_HEURES"), ("Besoin non résolu", "BESOIN_NON_RESOLU_SEANCES", "BESOIN_NON_RESOLU_HEURES")]
    heads = ["Hypothèse"]
    for lab, s, h in cols:
        heads += [lab + (" (séances)" if h else ""), lab + " (h)"] if h else [lab]
    header_row(bk, ws, r0 + 1, 1, heads, 45)
    for i in range(5):
        r = r0 + 2 + i
        bk.f(ws, r, 1, f'=INDEX(t_18[HYPOTHESE],{i + 1})&" : "&INDEX(t_18[DESCRIPTION],{i + 1})', "txt")
        c = 2
        for lab, s, h in cols:
            bk.f(ws, r, c, f"=INDEX(t_18[{s}],{i + 1})", "num")
            c += 1
            if h:
                bk.f(ws, r, c, f"=INDEX(t_18[{h}],{i + 1})", "hrs")
                c += 1
    last_col = 1 + len(heads) - 1
    ws.conditional_format(r0 + 2, 2, r0 + 6, last_col, {"type": "no_blanks", "format": bk.fmt["base"]})
    # --- réaffectation après chaque hypothèse ---
    r1 = r0 + 10
    bk.t(ws, r1, 1, "Réaffectation selon les effectifs, après chaque hypothèse", "h2")
    header_row(bk, ws, r1 + 1, 1, ["Situation", "Séances réaffectées", "Heures réaffectées", "Heures d'amphi libérées (net)", "Occupation amphis Pau : avant", "après", "Remplissage : avant", "après", "Effectif > capacité"], 45)
    for i, sit in enumerate(SITUATIONS):
        r = r1 + 2 + i
        bk.t(ws, r, 1, sit, "txt_b")
        k = realloc_kpis(sit, a1(r, 1, False))
        bk.f(ws, r, 2, k["moved_n"], "num")
        bk.f(ws, r, 3, k["moved_h"], "hrs")
        bk.f(ws, r, 4, f"={k['out_h'][1:]}-({k['in_h'][1:]})", "hrs")
        bk.f(ws, r, 5, k["occ_b"], "pct")
        bk.f(ws, r, 6, k["occ_a"], "pct")
        bk.f(ws, r, 7, k["fill_b"], "pct")
        bk.f(ws, r, 8, k["fill_a"], "pct")
        bk.f(ws, r, 9, k["over"], "num")
    # --- conflits avant optimisation ---
    r2 = r1 + 10
    bk.t(ws, r2, 1, "Conflits avant optimisation des locaux", "h2")
    bk.t(ws, r2 + 1, 1, "Résultat du report initial. Ce qui reste après relogement est dans l'onglet « Conclusion LET ».", "note")
    for i, key in enumerate(("H1a", "H1b", "H1c", "H2", "H3b")):
        r = r2 + 2 + i
        ws.merge_range(r, 1, r, 12, "", bk.fmt["txt"])
        bk.f(ws, r, 1, f'=IF({H("CHEVAUCHEMENTS", key)}>0,"{key} : "&{H("CHEVAUCHEMENTS", key)}&" moments où deux cours ont lieu en même temps dans la même salle.","{key} : aucun chevauchement dans les locaux d\'accueil.")', "txt")
    # --- données et limites ---
    r3 = r2 + 9
    bk.t(ws, r3, 1, "Données et limites de l'analyse", "h2")
    limits = [
        "Effectifs : EFFCALCU est un effectif calculé à partir des inscriptions (identifiants d'étudiants des groupes de la séance, groupes réunis cumulés), et non un effectif constaté en séance.",
        "Séances sans effectif : elles ne sont jamais déplacées et leurs heures restent comptées dans le besoin non résolu. Les effectifs estimés sont des majorants, présentés à part.",
        "Examens : examens, contrôles continus, rattrapages, oraux, soutenances et devoirs sont distingués des cours ordinaires ; ils sont exclus de la réaffectation globale (contraintes d'accueil propres).",
        "Période couverte : l'export s'arrête fin avril 2026 ; les examens de mai-juin ne sont pas couverts tant que l'export n'est pas complété jusqu'à la fin de l'année universitaire.",
        "Séances sans salle : une part importante des lignes de l'export n'a aucun code de salle ; l'occupation réelle des salles est donc sous-estimée.",
        "Réaffectation : périmètre = promotions « SSH - Pau », cours, CM, TD et CTD dans les locaux banalisés du campus de Pau, toutes capacités. Équipements, accessibilité et appartenance des locaux ne sont pas vérifiés.",
        "Lissage (H2, H3b) : test théorique de disponibilité ; les enseignants ne figurent pas dans l'export et ne sont pas contrôlés.",
    ]
    for i, t in enumerate(limits):
        ws.merge_range(r3 + 1 + i, 1, r3 + 1 + i, 12, t, bk.fmt["txt"])
        ws.set_row(r3 + 1 + i, 30)
    ws.freeze_panes(3, 0)
    ws.set_zoom(90)


def build_conclusion(bk: Book, info: dict):
    ws = bk.sheets["Conclusion LET"]
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 44)
    ws.set_column(2, 3, 16)
    ws.set_column(4, 4, 52)
    ws.set_column(5, 5, 60)
    ws.set_column(6, 24, 12)
    bk.t(ws, 0, 1, "Peut-on supprimer les trois amphis LET ?", "title")
    bk.t(ws, 1, 1, "Conclusion par scénario, lue dans les CSV 26 (après optimisation) et 27 (conclusion).", "sub")
    bk.t(ws, 3, 1, "Réponse par scénario", "h2")
    header_row(bk, ws, 4, 1, ["Scénario", "Horaires inchangés après optimisation", "Lissage complémentaire", "Conclusion", "Conditions à lever"], 45)
    n27 = len(info["tables"]["t_27"])
    for i in range(n27):
        r = 5 + i
        bk.f(ws, r, 1, f'=INDEX(t_27[key],{i + 1})&" : "&INDEX(t_27[label],{i + 1})', "txt")
        bk.f(ws, r, 2, f"=INDEX(t_27[horaires_inchanges],{i + 1})", "txt")
        bk.f(ws, r, 3, f"=INDEX(t_27[lissage],{i + 1})", "txt")
        bk.f(ws, r, 4, f"=INDEX(t_27[verdict],{i + 1})", "txt")
        bk.f(ws, r, 5, f"=INDEX(t_27[conditions],{i + 1})", "txt")
        ws.set_row(r, 62)
    rg = (5, 2, 4 + n27, 3)
    ws.conditional_format(*rg, {"type": "text", "criteria": "begins with", "value": "oui", "format": bk.fmt["good"]})
    ws.conditional_format(*rg, {"type": "text", "criteria": "begins with", "value": "non", "format": bk.fmt["bad"]})
    ws.conditional_format(5, 4, 4 + n27, 4, {"type": "text", "criteria": "begins with", "value": "Non démontré", "format": bk.fmt["bad"]})
    ws.conditional_format(5, 4, 4 + n27, 4, {"type": "text", "criteria": "begins with", "value": "Possible avec lissage", "format": bk.fmt["warn"]})
    ws.conditional_format(5, 4, 4 + n27, 4, {"type": "text", "criteria": "begins with", "value": "Possible à horaires", "format": bk.fmt["good"]})
    r0 = 5 + n27 + 2
    bk.t(ws, r0, 1, "Résultat après optimisation : chaque séance reportée est comptée une seule fois", "h2")
    cats = [("En conflit avant", "CONFLITS_AVANT"), ("Relogées pour résoudre un conflit", "RESOLUS_PAR_RELOGEMENT"), ("Placées sans conflit, effectif connu", "PLACEES_SANS_CONFLIT_EFFECTIF_CONNU"),
            ("Placées sans conflit, effectif estimé", "PLACEES_SANS_CONFLIT_EFFECTIF_ESTIME"), ("Restent en conflit (effectif connu)", "CONFLIT_PERSISTANT"),
            ("Restent en conflit (effectif incertain)", "CONFLIT_INCERTAIN"), ("Non placées : sans effectif", "NON_PLACEES_SANS_EFFECTIF"), ("Non placées : sans solution ou non couvertes", "NON_PLACEES_AUTRE")]
    ws.merge_range(r0 + 1, 2, r0 + 1, 3, "À reporter", bk.fmt["th"])
    for j, (lab, _) in enumerate(cats):
        ws.merge_range(r0 + 1, 4 + 2 * j, r0 + 1, 5 + 2 * j, lab, bk.fmt["th"])
    ws.set_row(r0 + 1, 45)
    bk.t(ws, r0 + 2, 1, "Hypothèse", "th")
    bk.t(ws, r0 + 2, 2, "Séances", "th")
    bk.t(ws, r0 + 2, 3, "Heures", "th")
    for j in range(len(cats)):
        bk.t(ws, r0 + 2, 4 + 2 * j, "Séances", "th")
        bk.t(ws, r0 + 2, 5 + 2 * j, "Heures", "th")
    for i in range(5):
        r = r0 + 3 + i
        bk.f(ws, r, 1, f"=INDEX(t_26[HYPOTHESE],{i + 1})", "txt_b")
        bk.f(ws, r, 2, f"=INDEX(t_26[SEANCES_A_REPORTER],{i + 1})", "num")
        bk.f(ws, r, 3, f"=INDEX(t_26[HEURES_A_REPORTER],{i + 1})", "hrs")
        for j, (_, pre) in enumerate(cats):
            bk.f(ws, r, 4 + 2 * j, f"=INDEX(t_26[{pre}_SEANCES],{i + 1})", "num")
            bk.f(ws, r, 5 + 2 * j, f"=INDEX(t_26[{pre}_HEURES],{i + 1})", "hrs")
    ws.set_column(4, 4, 18)
    ws.set_column(5, 5, 18)
    # vérification d'absence de double comptage : les 6 catégories exclusives somment au total à reporter
    r1 = r0 + 9
    bk.t(ws, r1, 1, "Contrôle : les 6 catégories exclusives (placées, en conflit, non placées) somment au total à reporter", "note")
    for i in range(5):
        r = r1 + 1 + i
        bk.f(ws, r, 1, f"=INDEX(t_26[HYPOTHESE],{i + 1})", "txt")
        excl = "+".join(f"INDEX(t_26[{pre}_SEANCES],{i + 1})" for pre in ("PLACEES_SANS_CONFLIT_EFFECTIF_CONNU", "PLACEES_SANS_CONFLIT_EFFECTIF_ESTIME", "CONFLIT_PERSISTANT", "CONFLIT_INCERTAIN", "NON_PLACEES_SANS_EFFECTIF", "NON_PLACEES_AUTRE"))
        bk.f(ws, r, 2, f"={excl}", "num")
        bk.f(ws, r, 3, f'=IF({a1(r, 2)}=INDEX(t_26[SEANCES_A_REPORTER],{i + 1}),"cohérent","ÉCART")', "txt")
    ws.conditional_format(r1 + 1, 3, r1 + 5, 3, {"type": "text", "criteria": "begins with", "value": "ÉCART", "format": bk.fmt["bad"]})
    ws.conditional_format(r1 + 1, 3, r1 + 5, 3, {"type": "text", "criteria": "begins with", "value": "cohérent", "format": bk.fmt["good"]})
    # conflits restants : mise en forme conditionnelle
    for pre_col in (8 + 4, ):
        pass
    cp = 4 + 2 * 4  # 1re colonne « Restent en conflit (effectif connu) » (séances)
    ws.conditional_format(r0 + 3, cp, r0 + 7, cp + 3, {"type": "cell", "criteria": ">", "value": 0, "format": bk.fmt["bad"]})
    ws.conditional_format(r0 + 3, cp, r0 + 7, cp + 3, {"type": "cell", "criteria": "==", "value": 0, "format": bk.fmt["good"]})
    r2 = r1 + 8
    bk.t(ws, r2, 1, "Moments en double qui restent", "h2")
    header_row(bk, ws, r2 + 1, 1, ["Hypothèse", "Moments (effectif connu)", "Heures doublées", "Moments (effectif incertain)", "Heures doublées", "Moments sans séance LET"], 45)
    for i in range(5):
        r = r2 + 2 + i
        bk.f(ws, r, 1, f"=INDEX(t_26[HYPOTHESE],{i + 1})", "txt_b")
        for j, (col, fm) in enumerate([("CONFLITS_PERSISTANTS_MOMENTS", "num"), ("CONFLITS_PERSISTANTS_HEURES_DOUBLEES", "hrs1"), ("CONFLITS_INCERTAINS_MOMENTS", "num"), ("CONFLITS_INCERTAINS_HEURES_DOUBLEES", "hrs1"), ("CONFLITS_SANS_SEANCE_LET", "num")]):
            bk.f(ws, r, 2 + j, f"=INDEX(t_26[{col}],{i + 1})", fm)
    ws.conditional_format(r2 + 2, 2, r2 + 6, 6, {"type": "cell", "criteria": ">", "value": 0, "format": bk.fmt["warn"]})
    # textes : démontré / à vérifier / conditions / validations
    r3 = r2 + 9
    T = lambda col, key: f'INDEX(t_26[{col}],MATCH("{key}",t_26[HYPOTHESE],0))'
    blocks = [
        ("Ce qui est démontré", [
            '="Les séances à effectif connu des trois amphis LET peuvent être relogées sur les locaux d\'accueil de chaque hypothèse : la capacité est vérifiée pour chacune."',
            f'="Après optimisation des locaux (relogement des séances en conflit dans un local libre de capacité suffisante, au même jour et à la même heure), il reste "&{T("CONFLIT_PERSISTANT_SEANCES", "H1a")}&" séances à effectif connu en conflit en H1a, "&{T("CONFLIT_PERSISTANT_SEANCES", "H1b")}&" en H1b et "&{T("CONFLIT_PERSISTANT_SEANCES", "H1c")}&" en H1c."',
            '="Le lissage (H2, H3b) supprime les chevauchements des séances à effectif connu en déplaçant des horaires, sans contrôle de la disponibilité des enseignants."']),
        ("Ce qui reste à vérifier", [
            f'="Séances LET placées sur effectif estimé : "&{T("PLACEES_SANS_CONFLIT_EFFECTIF_ESTIME_SEANCES", "H1a")}&" séances ("&ROUND({T("PLACEES_SANS_CONFLIT_EFFECTIF_ESTIME_HEURES", "H1a")},0)&" h) : effectif réel à confirmer (onglet À vérifier)."',
            f'="Usages sans effectif : "&{T("NON_PLACEES_SANS_EFFECTIF_SEANCES", "H1a")}&" séances ("&ROUND({T("NON_PLACEES_SANS_EFFECTIF_HEURES", "H1a")},0)&" h) à qualifier ; leurs heures restent à reloger (onglet Usages sans effectif)."',
            '="Séances sans salle dans l\'export et anomalies de capacité des locaux de report : voir l\'onglet À vérifier."',
            '="Examens de mai-juin : l\'export s\'arrête fin avril 2026 ; la conclusion reste assortie de cette réserve."']),
        ("Conditions nécessaires à la suppression", [
            '="Lever les besoins non pris en charge : confirmer l\'effectif, la capacité nécessaire ou l\'annulation des séances sans effectif ; tant que ce n\'est pas fait, leurs heures restent à reloger."',
            '="Pour toute séance restant en conflit : accepter un lissage ou un changement de local validé avec les composantes concernées."',
            '="Vérifier les séances placées sur effectif estimé, en commençant par celles de priorité haute."',
            '="Disposer d\'un export couvrant mai-juin pour vérifier la capacité d\'accueil des examens."']),
        ("Validations pratiques à conduire avec les gestionnaires", [
            '="Mutualisation des locaux : accord des composantes propriétaires pour l\'usage de leurs amphis et salles."',
            '="Équipements : sonorisation, visioprojection, enregistrement, accessibilité des locaux de report."',
            '="Déplacements : temps de trajet entre bâtiments pour les enseignants et les étudiants enchaînant deux cours."',
            '="Disponibilité des enseignants en cas de lissage (non contrôlée : le modèle ne dispose pas des enseignants)."',
            '="Capacité adaptée aux examens : espacement et surveillance que la capacité nominale ne représente pas."']),
    ]
    r = r3
    for title, items in blocks:
        bk.t(ws, r, 1, title, "h2")
        for line in items:
            r += 1
            ws.merge_range(r, 1, r, 8, "", bk.fmt["txt"])
            bk.f(ws, r, 1, line, "txt")
            ws.set_row(r, 30)
        r += 2
    ws.freeze_panes(3, 0)
    ws.set_zoom(90)


def build_verification(bk: Book, info: dict):
    ws = bk.sheets["À vérifier"]
    t = info["tables"]
    n23, n24, n25 = len(t["t_23"]), len(t["t_24"]), len(t["t_25"])
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 12, 16)
    ws.set_column(4, 5, 30)
    bk.t(ws, 0, 1, "Points à vérifier en priorité", "title")
    bk.t(ws, 1, 1, "Les trois sources d'incertitude qui peuvent changer la conclusion.", "sub")
    bk.t(ws, 3, 1, "1. Séances LET placées sur effectif estimé", "h2")
    header_row(bk, ws, 4, 1, ["Priorité", "Séances", "Heures (8 h-18 h)"], 30)
    for i, p in enumerate(["haute", "moyenne", "faible"]):
        r = 5 + i
        bk.t(ws, r, 1, p, "txt_b")
        bk.f(ws, r, 2, f"=COUNTIFS(t_23[PRIORITE],{a1(r, 1)})", "num")
        bk.f(ws, r, 3, f"=SUMIFS(t_23[HEURES_8H_18H],t_23[PRIORITE],{a1(r, 1)})", "hrs1")
    bk.t(ws, 8, 1, "Total", "txt_b")
    bk.f(ws, 8, 2, "=SUM(C6:C8)", "num")
    bk.f(ws, 8, 3, "=SUM(D6:D8)", "hrs1")
    header_row(bk, ws, 10, 1, ["Priorité", "Date", "Amphi LET", "Matière", "Groupe", "Effectif estimé", "Méthode", "Heures", "Local H1a (capacité)", "Pourquoi"])
    for i in range(1, n23 + 1):
        r = 10 + i
        g = lambda col: f"INDEX(t_23[{col}],{i})"
        for c, (expr, fm) in enumerate([(g("PRIORITE"), "txt"), (g("DATE"), "date"), (g("AMPHI_LET"), "txt"), (g("LIBELLE_MAT"), "txt"), (g("NOM_DIP"), "txt"), (g("EFFECTIF_ESTIME"), "num"),
                                       (g("METHODE"), "txt"), (g("HEURES_8H_18H"), "hrs1"), (f'{g("LOCAL_H1a")}&" ("&{g("CAPACITE_H1a")}&")"', "txt"), (g("POURQUOI"), "txt")]):
            bk.f(ws, r, 1 + c, f"={expr}", fm)
    ws.conditional_format(11, 1, 10 + n23, 1, {"type": "text", "criteria": "begins with", "value": "haute", "format": bk.fmt["bad"]})
    ws.conditional_format(11, 1, 10 + n23, 1, {"type": "text", "criteria": "begins with", "value": "moyenne", "format": bk.fmt["warn"]})
    r2 = 10 + n23 + 3
    bk.t(ws, r2, 1, "2. Séances SSH Pau sans salle dans l'export", "h2")
    header_row(bk, ws, r2 + 1, 1, ["Effectif du groupe", "Séries", "Séances", "Heures (8 h-18 h)", "dont examens", "Risque pour les amphis"], 30)
    brackets = [("Effectif inconnu", None, None, "à qualifier"), ("Moins de 50", 1, 49, "non"), ("50 à 99", 50, 99, "non"), ("100 à 149", 100, 149, "oui"), ("150 et plus", 150, 10**6, "oui")]
    for i, (lab, lo, hi, risk) in enumerate(brackets):
        r = r2 + 2 + i
        bk.t(ws, r, 1, lab, "txt_b")
        crit = 't_24[effect],""' if lo is None else f't_24[effect],">="&{lo},t_24[effect],"<="&{hi}'
        bk.f(ws, r, 2, f"=COUNTIFS({crit})", "num")
        bk.f(ws, r, 3, f"=SUMIFS(t_24[sessions],{crit})", "num")
        bk.f(ws, r, 4, f"=SUMIFS(t_24[hours],{crit})", "hrs")
        bk.f(ws, r, 5, f'=SUMIFS(t_24[sessions],{crit},t_24[exam],"True")', "num")
        bk.t(ws, r, 6, risk, "txt")
    ws.conditional_format(r2 + 2, 6, r2 + 6, 6, {"type": "text", "criteria": "begins with", "value": "oui", "format": bk.fmt["bad"]})
    ws.conditional_format(r2 + 2, 6, r2 + 6, 6, {"type": "text", "criteria": "begins with", "value": "à qualifier", "format": bk.fmt["warn"]})
    r3 = r2 + 9
    bk.t(ws, r3, 1, "Séries au plus grand effectif (40 premières)", "h2")
    header_row(bk, ws, r3 + 1, 1, ["Matière", "Type", "Groupes", "Jour", "Horaire", "Séances", "Effectif max"])
    for i in range(1, 41):
        r = r3 + 1 + i
        g = lambda col: f"INDEX(t_24[{col}],{i})"
        vals = [(g("subject"), "txt"), (g("type"), "txt"), (g("groups"), "txt"), (g("weekday"), "txt"),
                (f'INT({g("start")}/60)&"h"&TEXT(MOD({g("start")},60),"00")', "txt"), (g("sessions"), "num"), (f'IF({g("effect")}="","inconnu",{g("effect")})', "num")]
        for c, (expr, fm) in enumerate(vals):
            bk.f(ws, r, 1 + c, f"={expr}", fm)
    r4 = r3 + 45
    bk.t(ws, r4, 1, "3. Anomalies de capacité des locaux de report", "h2")
    header_row(bk, ws, r4 + 1, 1, ["Local", "Bât.", "Capacité", "Effectif max observé", "Anomalie", "Séances LET H1a (dont justes)", "Séances LET H1b (dont justes)", "Séances LET H1c (dont justes)"], 45)
    for i in range(1, n25 + 1):
        r = r4 + 1 + i
        g = lambda col: f"INDEX(t_25[{col}],{i})"
        vals = [(g("LOCAL"), "txt"), (g("BATIMENT"), "txt"), (g("CAPACITE"), "num"), (g("EFFECTIF_MAX_OBSERVE"), "num"), (g("ANOMALIE"), "txt")]
        vals += [(f'{g("SEANCES_LET_" + h)}&" ("&{g("DONT_JUSTES_" + h)}&")"', "txt") for h in ("H1a", "H1b", "H1c")]
        for c, (expr, fm) in enumerate(vals):
            bk.f(ws, r, 1 + c, f"={expr}", fm)
    ws.conditional_format(r4 + 2, 4, r4 + 1 + n25, 4, {"type": "cell", "criteria": ">", "value": 0, "format": bk.fmt["warn"]})
    ws.set_column(5, 5, 44)
    ws.freeze_panes(3, 0)
    ws.set_zoom(90)


def build_usages(bk: Book, info: dict):
    ws = bk.sheets["Usages sans effectif"]
    t17 = info["tables"]["t_17"]
    n17 = len(t17)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, 20, 14)
    bk.t(ws, 0, 1, "Usages sans effectif", "title")
    bk.t(ws, 1, 1, "Séances des amphis LET dont le planning ne donne pas l'effectif : qualifiées par usage d'après leurs intitulés ; leurs heures restent à reloger.", "sub")
    cards = [("Séances sans effectif", "=ROWS(t_17[ID_SEANCE])", "kpi_n"), ("Heures à reloger (8 h-18 h)", "=SUM(t_17[HEURES_8H_18H])", "kpi_h"),
             ("Heures, durée complète", "=SUM(t_17[HEURES_COMPLETES])", "kpi_h"), ("Part des séances à reporter", '=ROWS(t_17[ID_SEANCE])/INDEX(t_18[SEANCES_A_REPORTER],MATCH("H1a",t_18[HYPOTHESE],0))', "kpi_p")]
    for i, (label, formula, fm) in enumerate(cards):
        c = 1 + 2 * i if i else 1
        cc = [1, 3, 5, 7][i]
        ws.merge_range(3, cc, 3, cc + 1, label, bk.fmt["kpi_l"])
        ws.merge_range(4, cc, 4, cc + 1, "", bk.fmt[fm])
        bk.f(ws, 4, cc, formula, fm)
    ws.set_row(3, 30)
    ws.set_row(4, 28)
    cats = sorted(t17["CATEGORIE"].unique())
    rooms = sorted(t17["NOM_SAL"].unique())
    bk.t(ws, 6, 1, "Usages qualifiés", "h2")
    heads = ["Usage", "Séances", "Heures 8 h-18 h", "Heures durée complète"] + [f"Séances : {r}" for r in rooms]
    heads += [f"{h} : local libre ≥ {s} places" for h in ("H1a", "H1b", "H1c") for s in (50, 100, 150)]
    header_row(bk, ws, 7, 1, heads, 60)
    for i, cat in enumerate(cats):
        r = 8 + i
        bk.t(ws, r, 1, cat, "txt_b")
        k = a1(r, 1)
        bk.f(ws, r, 2, f"=COUNTIFS(t_17[CATEGORIE],{k})", "num")
        bk.f(ws, r, 3, f"=SUMIFS(t_17[HEURES_8H_18H],t_17[CATEGORIE],{k})", "hrs1")
        bk.f(ws, r, 4, f"=SUMIFS(t_17[HEURES_COMPLETES],t_17[CATEGORIE],{k})", "hrs1")
        for j, room in enumerate(rooms):
            bk.f(ws, r, 5 + j, f'=COUNTIFS(t_17[CATEGORIE],{k},t_17[NOM_SAL],"{room}")', "num")
        c = 5 + len(rooms)
        for h in ("H1a", "H1b", "H1c"):
            for s in (50, 100, 150):
                bk.f(ws, r, c, f'=COUNTIFS(t_17[CATEGORIE],{k},t_17[PLUS_GRAND_LOCAL_LIBRE_{h}],">="&{s})', "num")
                c += 1
    rt = 8 + len(cats)
    bk.t(ws, rt, 1, "Total", "txt_b")
    for c in range(2, 5 + len(rooms) + 9):
        col = xl_col_to_name(c)
        bk.f(ws, rt, c, f"=SUM({col}9:{col}{rt})", "num" if c not in (3, 4) else "hrs1")
    bar_chart(bk, ws, "column", "Heures à reloger par usage (8 h-18 h)", f"='Usages sans effectif'!{rng(8, 1, 7 + len(cats), 1)}", f"='Usages sans effectif'!{rng(8, 3, 7 + len(cats), 3)}", "N4", (520, 260), color=WARN, num_fmt="0")
    r2 = rt + 3
    bk.t(ws, r2, 1, "Liste des séances", "h2")
    heads = ["Date", "Début", "Fin", "Amphi", "Capacité", "Matière", "Groupe", "Usage", "Confiance", "Heures 8 h-18 h"]
    header_row(bk, ws, r2 + 1, 1, heads)
    for i in range(1, n17 + 1):
        r = r2 + 1 + i
        g = lambda col: f"INDEX(t_17[{col}],{i})"
        for c, (col, fm) in enumerate([("DATE", "date"), ("HDEBUT", "txt"), ("HFIN", "txt"), ("NOM_SAL", "txt"), ("CAPACITE_LOCAL", "num"), ("LIBELLE_MAT", "txt"), ("NOM_DIP", "txt"), ("CATEGORIE", "txt"), ("CONFIANCE", "txt"), ("HEURES_8H_18H", "hrs1")]):
            bk.f(ws, r, 1 + c, f"={g(col)}&\"\"" if fm == "txt" else f"={g(col)}", fm)
    ws.set_column(6, 7, 26)
    ws.freeze_panes(3, 0)
    ws.set_zoom(90)


def build_surdim(bk: Book, info: dict):
    ws = bk.sheets["Surdimensionnement"]
    df = info["tables"]["t_06"]
    ws.set_column(0, 0, 2)
    bk.t(ws, 0, 1, "Surdimensionnement des locaux", "title")
    bk.t(ws, 1, 1, "Séances (Cours, CM, TD, CTD) dont la salle offre au moins deux fois l'effectif et au moins 20 places inutilisées, avec la plus petite salle libre à toutes les occurrences de la série.", "sub")
    cards = [("Occurrences détectées", "=ROWS(t_06[date])", "kpi_n"), ("Places inutilisées (moyenne)", "=AVERAGE(t_06[unused])", "kpi_n"),
             ("Avec une salle de remplacement", "=COUNT(t_06[replacementCode])", "kpi_n"), ("Sans salle de remplacement", "=ROWS(t_06[date])-COUNT(t_06[replacementCode])", "kpi_n")]
    for i, (label, formula, fm) in enumerate(cards):
        cc = [1, 3, 5, 7][i]
        ws.merge_range(3, cc, 3, cc + 1, label, bk.fmt["kpi_l"])
        ws.merge_range(4, cc, 4, cc + 1, "", bk.fmt[fm])
        bk.f(ws, 4, cc, formula, fm)
    ws.set_row(3, 30)
    ws.set_row(4, 28)
    blds = sorted({str(b) for b in df["building"]})[:40]
    bk.t(ws, 6, 1, "Occurrences par bâtiment", "h2")
    header_row(bk, ws, 7, 1, ["Bâtiment", "Occurrences", "Places inutilisées (moy.)"], 30)
    for i, b in enumerate(blds):
        r = 8 + i
        bk.t(ws, r, 1, b, "txt_b")
        bk.f(ws, r, 2, f"=COUNTIFS(t_06[building],{a1(r, 1)})", "num")
        bk.f(ws, r, 3, f"=IFERROR(AVERAGEIFS(t_06[unused],t_06[building],{a1(r, 1)}),0)", "num1")
    bar_chart(bk, ws, "column", "Occurrences surdimensionnées par bâtiment", f"='Surdimensionnement'!{rng(8, 1, 7 + len(blds), 1)}", f"='Surdimensionnement'!{rng(8, 2, 7 + len(blds), 2)}", "G7", (640, 300))
    top = 8 + max(len(blds), 16) + 2
    bk.t(ws, top, 1, "Liste complète (tableau filtrable t_06, import du CSV 06)", "h2")
    write_table(bk, "Surdimensionnement", "t_06", df, [], {"room": 34, "subject": 30, "diplomas": 30, "replacement": 34}, ws=ws, start_row=top + 1)
    ws.freeze_panes(3, 0)
    ws.set_zoom(90)


def build_sources(bk: Book, info: dict, src: Sources):
    ws = bk.sheets["Sources"]
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 52)
    ws.set_column(2, 3, 14)
    ws.set_column(4, 4, 20)
    ws.set_column(5, 5, 60)
    ws.set_column(6, 6, 60)
    bk.t(ws, 0, 1, "Sources", "title")
    bk.t(ws, 1, 1, f"Classeur généré le {dt.datetime.now():%d/%m/%Y à %H:%M} à partir du dossier {os.path.basename(src.folder)} (script build_dashboard_xlsx.py).", "sub")
    header_row(bk, ws, 3, 1, ["CSV utilisé", "Lignes (CSV)", "Lignes (classeur)", "Date de génération", "Feuilles alimentées", "Remarque"], 30)
    tables = {"06": ("t_06", None), "07": ("t_07", None), "09": ("t_09", None), "10": ("t_10", None), "11": ("t_11", None), "12": ("t_12", None), "13": ("t_13", None), "14": ("t_14", None),
              "17": ("t_17", None), "18": ("t_18", None), "23": ("t_23", None), "24": ("t_24", None), "25": ("t_25", None), "26": ("t_26", None), "27": ("t_27", None), "28": ("t_28", None)}
    for i, (name, meta) in enumerate(sorted(src.used.items())):
        r = 4 + i
        bk.t(ws, r, 1, name, "txt")
        ws.write_number(r, 2, meta["rows"], bk.fmt["num"])
        num = name[:2]
        m = re.search(r"_(reference|H1a|H1b|H1c|H2|H3b)_\d", name)
        if m and num in ("19", "20", "21", "22"):
            sit = "Référence" if m.group(1) == "reference" else m.group(1)
            bk.f(ws, r, 3, f'=COUNTIFS(t_{num}[SITUATION],"{sit}")', "num")
        elif num == "12" and "examens" in name:
            bk.f(ws, r, 3, "=ROWS(t_12X[date])", "num")
        elif num in tables:
            bk.f(ws, r, 3, f"=ROWS({tables[num][0]}[{info['first_col'][tables[num][0]]}])", "num")
        else:
            bk.t(ws, r, 3, "", "txt")
        ws.write_datetime(r, 4, meta["date"], bk.fmt["txt"] if False else bk.wb.add_format({"num_format": "dd/mm/yyyy hh:mm", "border": 1, "border_color": LINE, "font_size": 10}))
        bk.t(ws, r, 5, meta["sheets"], "txt")
        bk.t(ws, r, 6, meta["note"], "txt")
    last = 4 + len(src.used)
    ws.conditional_format(4, 3, last, 3, {"type": "formula", "criteria": f"=$D5<>$C5", "format": bk.fmt["bad"]})
    r = last + 2
    bk.t(ws, r, 1, "Feuilles techniques", "h2")
    notes = [("Listes", "Listes déroulantes, jours pédagogiques (déduits des heures/taux du CSV 28), catalogue des salles lu par formule."),
             ("Calcul", "Minutes d'occupation de chaque salle et de chaque tranche horaire, par SOMME.SI.ENS sur Données_07."),
             ("Données_Report", "Mise à plat des colonnes <hypothèse>_STATUT/_CODE/_DATE/_DEBUT du CSV 07 (une ligne par séance LET placée et par hypothèse), avec colonnes calculées."),
             ("Données_xx", "Import brut des CSV en tableaux structurés Txx. Les CSV 19 à 22 (un par situation) sont empilés avec une colonne SITUATION.")]
    for i, (a, b) in enumerate(notes):
        bk.t(ws, r + 1 + i, 1, a, "txt_b")
        ws.merge_range(r + 1 + i, 2, r + 1 + i, 6, b, bk.fmt["txt"])
    ws.freeze_panes(4, 0)


# =============================================================================================
# 7. Assemblage
# =============================================================================================
def build(folder: str, out: str, cache: dict | None = None):
    src = Sources(folder)
    bk = Book(out)
    bk.cache = cache or {}
    order = ["Conclusion LET", "Synthèse V2", "Amphis", "Réaffectation globale", "À vérifier", "Usages sans effectif"] + [VIEW_SHEETS[s] for s in SITUATIONS] + ["Surdimensionnement"]
    for name in order:
        bk.sheet(name, OK)
    bk.sheet("Sources", "#52636b")
    bk.sheet("Listes", "#52636b", hide_grid=False)
    bk.sheet("Calcul", "#52636b", hide_grid=False)
    info = build_data_sheets(bk, src)
    info["first_col"] = {"t_06": "date", "t_07": "ID_SEANCE", "t_09": "code", "t_10": "perimetre", "t_11": "perimetre", "t_12": "date", "t_13": "scenario", "t_14": "type",
                         "t_17": "ID_SEANCE", "t_18": "HYPOTHESE", "t_23": "ID_SEANCE", "t_24": "subject", "t_25": "LOCAL", "t_26": "HYPOTHESE", "t_27": "key", "t_28": "SITUATION"}
    only = [x for x in os.environ.get("XLSX_SEULEMENT", "").split(",") if x]  # mise au point : limiter les feuilles construites
    run = lambda name: not only or name in only
    build_listes(bk, info)
    build_calcul(bk, info)
    for name, fn in (("conclusion", build_conclusion), ("synthese", build_synthese), ("amphis", build_amphis), ("realloc", build_realloc),
                     ("verification", build_verification), ("usages", build_usages)):
        if run(name):
            fn(bk, info)
    for scen in SITUATIONS:
        if run("vues"):
            build_view(bk, info, scen, VIEW_SHEETS[scen])
    if run("surdim"):
        build_surdim(bk, info)
    build_sources(bk, info, src)
    bk.sheets["Conclusion LET"].activate()
    bk.wb.close()
    return info


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("dossier", help="Dossier resultats_AAAAMMJJ_HHMMSS_xxxxxx contenant les CSV")
    ap.add_argument("--sortie", help="Chemin du .xlsx (par défaut : 08_dashboard_complet_<horodatage>.xlsx dans le dossier)")
    ap.add_argument("--cache", help="JSON de valeurs calculées (feuille!cellule → valeur) à embarquer comme valeurs en cache")
    args = ap.parse_args()
    folder = os.path.abspath(args.dossier)
    stamp = re.search(r"resultats_(\d{8}_\d{6}_\d+)", folder)
    out = args.sortie or os.path.join(folder, f"08_dashboard_complet_{stamp.group(1) if stamp else 'xx'}.xlsx")
    cache = {}
    if args.cache:
        raw = json.load(open(args.cache, encoding="utf-8"))
        cache = {tuple(k.split("!", 1)): v for k, v in raw.items()}
    build(folder, out, cache)
    print(f"Classeur écrit : {out}")


if __name__ == "__main__":
    main()
