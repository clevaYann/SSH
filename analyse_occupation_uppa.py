#!/usr/bin/env python3
"""Précontrôle des sources pour l'analyse d'occupation UPPA.

Ce programme ne calcule pas d'occupation si le planning ne fournit pas les
occurrences nécessaires. Il produit alors une page HTML de diagnostic plutôt
que des résultats qui sembleraient complets.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import html
import io
import json
import math
import re
import sys
import tempfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any


SCHEDULE_REQUIRED = {
    "NOM_DIP",
    "TYPE",
    "PERIODE",
    "JOUR",
    "HDEBUT",
    "HFIN",
    "DUREE",
    "CODE_SAL",
}
ROOM_REQUIRED = {"CODE", "Nom", "CAPA"}
PROMOTION_REQUIRED = {"NOM_DIP", "EFFCALCU", "FAMILLE"}
CSV_SUFFIX = ".csv"


class SourceError(Exception):
    """Erreur de lecture d'une source tabulaire."""


def safe_text(value: Any) -> str:
    return html.escape(str(value if value is not None else ""), quote=True)


def read_csv(path: Path) -> dict[str, Any]:
    raw = path.read_bytes()
    text = None
    encoding = ""
    for candidate in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(candidate)
            encoding = candidate
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise SourceError(f"Encodage illisible : {path}")

    first_line = next((line for line in text.splitlines() if line.strip()), "")
    if not first_line:
        raise SourceError(f"Fichier vide : {path}")
    delimiter = max((";", ",", "\t", "|"), key=first_line.count)
    reader = csv.DictReader(io.StringIO(text, newline=""), delimiter=delimiter)
    original_fieldnames = reader.fieldnames or []
    occurrences: Counter[str] = Counter()
    fieldnames: list[str] = []
    for fieldname in original_fieldnames:
        occurrences[fieldname] += 1
        fieldnames.append(
            fieldname if occurrences[fieldname] == 1 else f"{fieldname}__{occurrences[fieldname]}"
        )
    reader.fieldnames = fieldnames
    rows = list(reader)
    return {
        "path": path,
        "encoding": encoding,
        "delimiter": delimiter,
        "columns": fieldnames,
        "rows": rows,
        "bytes": len(raw),
    }


def find_source(directory: Path, explicit: str | None, patterns: tuple[str, ...]) -> Path | None:
    if explicit:
        path = Path(explicit).expanduser().resolve()
        return path if path.is_file() else None
    candidates: list[Path] = []
    for pattern in patterns:
        candidates.extend(directory.glob(pattern))
    unique = sorted(set(candidates), key=lambda item: item.name.casefold())
    return unique[0] if len(unique) == 1 else None


def blank_percentages(source: dict[str, Any]) -> dict[str, float]:
    rows = source["rows"]
    if not rows:
        return {column: 0.0 for column in source["columns"]}
    return {
        column: 100.0
        * sum(not str(row.get(column) or "").strip() for row in rows)
        / len(rows)
        for column in source["columns"]
    }


def parse_decimal(value: str) -> float | None:
    try:
        return float(value.strip().replace(",", "."))
    except (AttributeError, ValueError):
        return None


def get_audit_facts(source: dict[str, Any], kind: str) -> list[str]:
    rows = source["rows"]
    facts = [f"{len(rows):,} lignes de données".replace(",", " ")]
    if kind == "planning":
        if "DDEBUT" in source["columns"]:
            dates: list[dt.date] = []
            invalid_dates = 0
            mismatched_days = 0
            weekdays = ("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")
            for row in rows:
                value = str(row.get("DDEBUT") or "").strip()
                if not value:
                    continue
                parsed = None
                for format_string in ("%d/%m/%Y", "%Y-%m-%d"):
                    try:
                        parsed = dt.datetime.strptime(value, format_string).date()
                        break
                    except ValueError:
                        pass
                if parsed is None:
                    invalid_dates += 1
                    continue
                dates.append(parsed)
                day = str(row.get("JOUR") or "").strip().casefold()
                if day and day != weekdays[parsed.weekday()]:
                    mismatched_days += 1
            if dates:
                facts.append(f"DDEBUT : {min(dates).isoformat()} à {max(dates).isoformat()}")
            facts.append(f"Dates illisibles : {invalid_dates} ; jours incohérents : {mismatched_days}")

        durations = [
            parsed
            for row in rows
            if (parsed := parse_decimal(str(row.get("DUREE") or ""))) is not None
        ]
        if durations:
            facts.append(
                "DUREE : minimum {} h, maximum {} h, valeurs >= 8 h : {}".format(
                    min(durations),
                    max(durations),
                    sum(duration >= 8 for duration in durations),
                )
            )
        if "NOM_SAL" in source["columns"]:
            multi_room = sum("@" in str(row.get("NOM_SAL") or "") for row in rows)
            facts.append(f"Lignes NOM_SAL contenant @ : {multi_room}")
    elif kind == "rooms":
        codes = [str(row.get("CODE") or "").strip() for row in rows]
        duplicate_codes = sum(count - 1 for count in Counter(codes).values() if count > 1)
        zero_capacity = sum(parse_decimal(str(row.get("CAPA") or "")) == 0 for row in rows)
        facts.extend(
            (
                f"CODE vide : {sum(not code for code in codes)}",
                f"Lignes avec CODE répété au-delà de la première : {duplicate_codes}",
                f"Capacités égales à zéro : {zero_capacity}",
            )
        )
    elif kind == "promotions":
        duplicate_names = sum(
            count - 1
            for count in Counter(str(row.get("NOM_DIP") or "").strip() for row in rows).values()
            if count > 1
        )
        facts.append(f"Lignes NOM_DIP répétées au-delà de la première : {duplicate_names}")
    return facts


def make_effects_template(planning: dict[str, Any], destination: Path) -> int:
    if "NOM_DIP" not in planning["columns"]:
        raise SourceError("Impossible de créer le modèle : la colonne NOM_DIP est absente du planning.")
    names = sorted(
        {
            str(row.get("NOM_DIP") or "").strip()
            for row in planning["rows"]
            if str(row.get("NOM_DIP") or "").strip()
        },
        key=str.casefold,
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream, delimiter=";", lineterminator="\n")
        writer.writerow(("NOM_DIP", "EFFECTIF_MANUEL"))
        writer.writerows((name, "") for name in names)
    return len(names)


def create_run_directory(parent: Path) -> tuple[Path, str, str]:
    created = dt.datetime.now().astimezone()
    stamp = created.strftime("%Y%m%d_%H%M%S_%f")
    run_directory = parent / f"resultats_{stamp}"
    suffix = 1
    while run_directory.exists():
        run_directory = parent / f"resultats_{stamp}_{suffix}"
        suffix += 1
    run_directory.mkdir(parents=True)
    return run_directory, stamp, created.isoformat(timespec="seconds")


def render_diagnostic(
    destination: Path,
    created_at: str,
    sources: dict[str, dict[str, Any]],
    blockers: list[str],
    arguments: argparse.Namespace,
) -> None:
    source_sections: list[str] = []
    for kind, source in sources.items():
        title = {"planning": "Planning", "rooms": "Locaux", "promotions": "Effectifs de référence"}[kind]
        percentages = blank_percentages(source)
        rows = "".join(
            "<tr><td>{}</td><td>{:.2f} %</td></tr>".format(safe_text(column), percentages[column])
            for column in source["columns"]
        )
        facts = "".join(f"<li>{safe_text(fact)}</li>" for fact in get_audit_facts(source, kind))
        source_sections.append(
            f"<section><h2>{safe_text(title)}</h2>"
            f"<p><strong>Fichier :</strong> {safe_text(source['path'].name)} · "
            f"{len(source['rows']):,} lignes · encodage détecté : {safe_text(source['encoding'])} · "
            f"séparateur : {safe_text(repr(source['delimiter']))}</p>".replace(",", " ")
            + f"<ul>{facts}</ul><details><summary>Valeurs vides par colonne</summary>"
            f"<table><thead><tr><th>Colonne</th><th>Vides</th></tr></thead><tbody>{rows}</tbody></table></details></section>"
        )

    blocker_items = "".join(f"<li>{safe_text(blocker)}</li>" for blocker in blockers)
    range_text = f"{arguments.debut or 'non précisé'} → {arguments.fin or 'non précisé'}"
    exclusions = arguments.exclure_semaines or "aucune"
    document = f"""<!doctype html>
<html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Précontrôle occupation UPPA — {safe_text(created_at)}</title>
<style>
:root{{color-scheme:light;--ink:#19252a;--muted:#52636b;--line:#d3dcdf;--paper:#f3f6f4;--white:#fff;--alert:#8e3029;--accent:#17675d}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--paper);color:var(--ink);font:16px/1.5 Georgia,serif}}
main{{max-width:980px;margin:32px auto;padding:0 22px 48px}}header{{border-bottom:2px solid var(--ink);padding:18px 0 24px}}
.eyebrow{{font:600 12px/1.4 monospace;text-transform:uppercase;color:var(--accent)}}h1{{font-size:clamp(27px,5vw,42px);line-height:1.1;margin:10px 0}}
header p,.meta{{color:var(--muted)}}section{{padding:20px 0;border-bottom:1px solid var(--line)}}h2{{font-size:22px;margin:0 0 12px}}
.alert{{border-left:5px solid var(--alert);padding:16px 18px;background:#fff0ec}}.alert h2{{color:var(--alert)}}li{{margin:5px 0}}
table{{width:100%;border-collapse:collapse;background:var(--white);font:14px/1.45 system-ui,sans-serif}}th,td{{text-align:left;padding:7px 9px;border-bottom:1px solid var(--line)}}th{{background:#e7eeee}}
details{{margin-top:10px}}summary{{cursor:pointer;color:var(--accent);font-family:system-ui,sans-serif}}code{{font-family:monospace;background:#e7eeee;padding:2px 4px}}
@media(max-width:600px){{main{{margin-top:14px;padding-inline:14px}}}}
</style></head><body><main>
<header><p class="eyebrow">UPPA · Direction du Patrimoine · SSH</p><h1>Précontrôle des données</h1>
<p>Créé le {safe_text(created_at)} · Période demandée : {safe_text(range_text)} · Semaines exclues : {safe_text(exclusions)}</p></header>
<section class="alert"><h2>Calcul interrompu, aucun indicateur publié</h2>
<p>Les données présentes ne permettent pas de reconstituer le calendrier ni d'identifier de façon fiable les locaux. Aucun taux, scénario ou résultat de faisabilité n'a été calculé.</p>
<ul>{blocker_items}</ul></section>
<section><h2>Entrées nécessaires</h2><ul>
<li>Un planning avec <code>PERIODE</code> et <code>CODE_SAL</code>, ou une spécification validée qui autorise explicitement une autre méthode d'occurrences et d'identification des salles.</li>
<li>La convention exacte de rapprochement entre <code>NOM_DIP</code> du planning et <code>EFFCALCU</code> des exports de promotions, y compris les lignes/familles à agréger, ou un fichier d'effectifs manuels validé.</li>
<li>Le calendrier pédagogique officiel (jours sans cours et périodes de fermeture) ou la liste à fournir avec <code>--exclure-semaines</code>. Les seuls jours fériés nationaux ne déterminent pas les vacances universitaires.</li>
<li>Le contenu lisible et validé de la méthodologie et des consignes complémentaires. Le compte rendu DOCX fourni est vide.</li>
</ul><p>Entrées CLI acceptées par ce précontrôle : <code>--dossier</code>, <code>--planning</code>, <code>--salles</code>, <code>--promotions</code>, <code>--debut</code>, <code>--fin</code>, <code>--exclure-semaines</code>. <code>--modele-effectifs</code> écrit les noms distincts de <code>NOM_DIP</code> pour saisie manuelle ; ces noms ne sont pas présumés être des groupes de présence.</p></section>
{''.join(source_sections)}
<footer><p>Cette page est un diagnostic de sources, pas le rapport d'occupation demandé. Elle ne contient aucune estimation.</p></footer>
</main></body></html>"""
    destination.write_text(document, encoding="utf-8")


class AnalysisError(Exception):
    """Erreur bloquante de données ou d'invariant de simulation."""


@dataclass(frozen=True)
class Room:
    code: str
    name: str
    capacity: int | None
    building: str
    kind: str


def iter_csv_dicts(path: Path):
    with path.open("rb") as stream:
        sample = stream.read(8192)
    encoding = ""
    for candidate in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            sample.decode(candidate)
            encoding = candidate
            break
        except UnicodeDecodeError:
            continue
    if not encoding:
        raise SourceError(f"Encodage illisible : {path}")
    with path.open("r", encoding=encoding, newline="") as stream:
        first = next((line for line in stream if line.strip()), "")
        if not first:
            raise SourceError(f"Fichier vide : {path}")
        delimiter = max((";", ",", "\t", "|"), key=first.count)
        stream.seek(0)
        reader = csv.reader(stream, delimiter=delimiter)
        original_headers = next(reader, [])
        counts: Counter[str] = Counter()
        headers: list[str] = []
        for header in original_headers:
            counts[header] += 1
            headers.append(header if counts[header] == 1 else f"{header}__{counts[header]}")
        yield {"__headers__": headers, "__encoding__": encoding, "__delimiter__": delimiter}
        for line_number, values in enumerate(reader, start=2):
            if len(values) != len(headers):
                raise SourceError(
                    f"{path.name}, ligne {line_number} : {len(values)} champs au lieu de {len(headers)}."
                )
            yield dict(zip(headers, values))


def parse_period_weeks(value: str) -> list[int]:
    match = re.fullmatch(r"\s*\[\s*(.*?)\s*\]\s*", value)
    if not match:
        raise SourceError(f"PERIODE illisible : {value!r} (format attendu : [37], [37..43] ou [52..1]).")
    weeks: list[int] = []
    for token in match.group(1).split(","):
        token = token.strip()
        if ".." in token:
            parts = token.split("..")
            if len(parts) != 2 or not all(part.strip().isdigit() for part in parts):
                raise SourceError(f"Plage PERIODE illisible : {value!r}.")
            start, end = (int(part.strip()) for part in parts)
            if start < 1 or start > 53 or end < 1 or end > 53:
                raise SourceError(f"Semaine ISO hors limites dans PERIODE : {value!r}.")
            if end >= start:
                weeks.extend(range(start, end + 1))
            else:
                weeks.extend(range(start, 54))
                weeks.extend(range(1, end + 1))
        elif token.isdigit() and 1 <= int(token) <= 53:
            weeks.append(int(token))
        else:
            raise SourceError(f"Semaine ISO illisible dans PERIODE : {value!r}.")
    if not weeks:
        raise SourceError(f"PERIODE vide : {value!r}.")
    return list(dict.fromkeys(weeks))


def dates_from_period(value: str, anchor: dt.date) -> list[dt.date]:
    weeks = parse_period_weeks(value)
    anchor_iso = anchor.isocalendar()
    if anchor_iso.week not in weeks:
        raise SourceError(
            f"PERIODE {value!r} ne contient pas la semaine ISO de DDEBUT {anchor.isoformat()}."
        )
    dates: list[dt.date] = []
    for week in weeks:
        year = anchor_iso.year
        if anchor_iso.week - week > 26:
            year += 1
        elif week - anchor_iso.week > 26:
            year -= 1
        try:
            occurrence = dt.date.fromisocalendar(year, week, anchor_iso.weekday)
        except ValueError:
            continue
        dates.append(anchor if occurrence == anchor else occurrence)
    return sorted(set(dates))


def parse_clock(value: str) -> int:
    match = re.fullmatch(r"\s*(\d{1,2}):(\d{2})\s*", value)
    if not match:
        raise SourceError(f"Heure illisible : {value!r}.")
    hour, minute = (int(part) for part in match.groups())
    if hour > 23 or minute > 59:
        raise SourceError(f"Heure hors limites : {value!r}.")
    return hour * 60 + minute


def building_from_name(name: str) -> str:
    matches = re.findall(r"\(([^()]*)\)", name)
    return matches[-1].split("-", 1)[0].strip() if matches else ""


# Codes de bâtiment d'EXP_SALLE absents ou écrits autrement dans SURFACES GENERAL.
BUILDING_ALIASES = {
    "B45": "B4B5",
    "DAL_R": "DAL",
    "SCI_R": "SCI",
    "IPM": "IPM1",
    "XPL": "XLP",
    "GTR": "RT",
}
EXTRA_BUILDING_SITES = {"TEL": ("TARBES", "TÉLÉSITE TARBES")}


def load_building_sites(path: Path) -> dict[str, dict[str, str]]:
    """Retourne {code bâtiment EXP_SALLE: {ville, site, nom}} depuis SURFACES GENERAL."""
    try:
        import openpyxl
    except ImportError as error:
        raise SourceError("Le module openpyxl est requis pour lire le fichier SURFACES (pip install openpyxl).") from error
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook["GENERAL SURFACES"]
        rows = sheet.iter_rows(min_row=2, max_col=4, values_only=True)
        sites = {
            str(code).strip(): {"ville": str(city).strip(), "site": str(site).strip(), "nom": str(name or "").strip()}
            for city, site, code, name in rows
            if city and site and code
        }
    finally:
        workbook.close()
    for alias, target in BUILDING_ALIASES.items():
        if target in sites:
            sites.setdefault(alias, sites[target])
    for code, (city, site) in EXTRA_BUILDING_SITES.items():
        sites.setdefault(code, {"ville": city, "site": site, "nom": site})
    return sites


def classify_room(name: str, capacity: int | None, building: str) -> str:
    lowered = name.casefold()
    if "amphi" in lowered:
        return "amphi"
    if "salle de cours" in lowered:
        if capacity is not None and 50 <= capacity <= 100 and building in {"LET", "DEG"}:
            return "grande_salle"
        return "salle_cours"
    return "autre"


def load_rooms(path: Path) -> tuple[dict[str, Room], list[str]]:
    rows = iter_csv_dicts(path)
    metadata = next(rows)
    headers = set(metadata["__headers__"])
    missing = ROOM_REQUIRED - headers
    if missing:
        raise SourceError(f"Colonnes manquantes dans {path.name} : {', '.join(sorted(missing))}.")
    rooms: dict[str, Room] = {}
    duplicate_codes: Counter[str] = Counter()
    for row in rows:
        code = row.get("CODE", "").strip()
        name = row.get("Nom", "").strip()
        capacity_text = row.get("CAPA", "").strip()
        if not code:
            continue
        try:
            capacity = int(capacity_text)
        except ValueError as error:
            raise SourceError(f"Capacité non entière dans {path.name} : {name!r}, {capacity_text!r}.") from error
        building = building_from_name(name)
        room = Room(code, name, capacity, building, classify_room(name, capacity, building))
        if code in rooms:
            duplicate_codes[code] += 1
            if rooms[code] != room:
                raise SourceError(f"CODE de local {code} associé à plusieurs noms ou capacités.")
        else:
            rooms[code] = room
    return rooms, [f"Codes répétés identiques ignorés : {sum(duplicate_codes.values())}"]


def easter_sunday(year: int) -> dt.date:
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = (h + l - 7 * m + 114) % 31 + 1
    return dt.date(year, month, day)


def french_public_holidays(year: int) -> set[dt.date]:
    easter = easter_sunday(year)
    return {
        dt.date(year, 1, 1),
        easter + dt.timedelta(days=1),
        dt.date(year, 5, 1),
        dt.date(year, 5, 8),
        easter + dt.timedelta(days=39),
        easter + dt.timedelta(days=50),
        dt.date(year, 7, 14),
        dt.date(year, 8, 15),
        dt.date(year, 11, 1),
        dt.date(year, 11, 11),
        dt.date(year, 12, 25),
    }


def parse_effect(value: str) -> int | None:
    parsed = parse_decimal(value)
    if parsed is None or parsed <= 0 or not parsed.is_integer():
        return None
    return int(parsed)


def load_schedule(
    path: Path,
    rooms: dict[str, Room],
    promotion_names: set[str],
    start: dt.date | None,
    end: dt.date | None,
    excluded_weeks: set[int],
    closed_dates: set[dt.date] | None = None,
) -> dict[str, Any]:
    closed_dates = closed_dates or set()
    iterator = iter_csv_dicts(path)
    metadata = next(iterator)
    headers = metadata["__headers__"]
    missing = (SCHEDULE_REQUIRED | {"DDEBUT", "HFIN", "LIBELLE_MAT", "EFFCALCU"}) - set(headers)
    if missing:
        raise SourceError(f"Colonnes manquantes dans {path.name} : {', '.join(sorted(missing))}.")

    blank_counts: Counter[str] = Counter()
    raw_rows = 0
    date_min: dt.date | None = None
    date_max: dt.date | None = None
    iso_mismatch = 0
    day_mismatch = 0
    duration_mismatch = 0
    at_room_count = 0
    at_name_count = 0
    at_code_count = 0
    period_values: set[str] = set()
    period_range_rows = 0
    duration_min: float | None = None
    duration_max: float | None = None
    admin_count = 0
    holiday_count = 0
    closure_count = 0
    weekend_count = 0
    excluded_week_count = 0
    outside_range_count = 0
    missing_room_code_count = 0
    unknown_codes: Counter[str] = Counter()
    prefixed_normalized = 0
    prefixed_rejected: Counter[str] = Counter()
    schedule_dates: set[dt.date] = set()
    unrecognized_promotions: set[str] = set()
    grouped: dict[tuple[Any, ...], dict[str, Any]] = {}
    weekday_names = ("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")
    holidays = french_public_holidays(2025) | french_public_holidays(2026)

    for row in iterator:
        raw_rows += 1
        for field in headers:
            if not str(row.get(field, "")).strip():
                blank_counts[field] += 1
        diploma = str(row.get("NOM_DIP", "")).strip()
        if diploma and diploma not in promotion_names:
            unrecognized_promotions.add(diploma)
        room_name = str(row.get("NOM_SAL", "")).strip()
        room_code = str(row.get("CODE_SAL", "")).strip()
        if room_code.startswith("<>"):
            # Code préfixé « <>NNNN » avec un nom « <Groupe>Nom (BAT-00) » : même salle que NNNN si le nom concorde avec EXP_SALLE.
            bare = room_code[2:]
            plain_name = re.sub(r"^<[^>]*>", "", room_name).strip()
            if bare in rooms and rooms[bare].name == plain_name:
                prefixed_normalized += 1
                room_code, room_name = bare, plain_name
            else:
                prefixed_rejected[room_code] += 1
        if "@" in room_name:
            at_name_count += 1
        if "@" in room_code:
            at_code_count += 1
        if "@" in room_name or "@" in room_code:
            at_room_count += 1
        if room_name and not room_code:
            missing_room_code_count += 1
        if room_code and room_code not in rooms:
            unknown_codes[room_code] += 1

        try:
            anchor = dt.datetime.strptime(str(row.get("DDEBUT", "")).strip(), "%d/%m/%Y").date()
        except ValueError as error:
            raise SourceError(f"Date DDEBUT invalide à la ligne {raw_rows + 1} : {row.get('DDEBUT')!r}.") from error
        date_min = anchor if date_min is None or anchor < date_min else date_min
        date_max = anchor if date_max is None or anchor > date_max else date_max
        period_value = str(row.get("PERIODE", "")).strip()
        period_values.add(period_value)
        period_weeks = parse_period_weeks(period_value)
        periods = dates_from_period(period_value, anchor)
        if len(periods) > 1:
            period_range_rows += 1
        if anchor.isocalendar().week not in period_weeks:
            iso_mismatch += 1
        day = str(row.get("JOUR", "")).strip().casefold()
        if day and day != weekday_names[anchor.weekday()]:
            day_mismatch += 1

        start_minute = parse_clock(str(row.get("HDEBUT", "")).strip())
        end_minute = parse_clock(str(row.get("HFIN", "")).strip())
        if end_minute < start_minute:
            end_minute += 24 * 60
        duration = end_minute - start_minute
        declared_duration = parse_decimal(str(row.get("DUREE", "")))
        if declared_duration is not None:
            duration_min = declared_duration if duration_min is None else min(duration_min, declared_duration)
            duration_max = declared_duration if duration_max is None else max(duration_max, declared_duration)
        if declared_duration is not None and abs(declared_duration * 60 - duration) > 1.0:
            duration_mismatch += 1
        if duration <= 0:
            raise SourceError(f"Durée nulle ou négative à la ligne {raw_rows + 1}.")
        if duration >= 8 * 60:
            admin_count += 1
            continue

        subject = str(row.get("LIBELLE_MAT", "")).strip()
        course_type = str(row.get("TYPE", "")).strip()
        effect = parse_effect(str(row.get("EFFCALCU", "")).strip())
        raw_effect = str(row.get("EFFCALCU", "")).strip()
        for occurrence in periods:
            if start and occurrence < start or end and occurrence > end:
                outside_range_count += 1
                continue
            if occurrence in holidays:
                holiday_count += 1
                continue
            if occurrence in closed_dates:
                closure_count += 1
                continue
            if occurrence.isoweekday() > 5:
                weekend_count += 1
                continue
            if occurrence.isocalendar().week in excluded_weeks:
                excluded_week_count += 1
                continue
            schedule_dates.add(occurrence)
            if not room_code:
                continue
            room = rooms.get(room_code)
            actual_name = room.name if room else room_name
            building = room.building if room else building_from_name(actual_name)
            capacity = room.capacity if room else None
            kind = room.kind if room else classify_room(actual_name, capacity, building)
            identity = (occurrence, room_code, start_minute, end_minute, subject, course_type)
            entry = grouped.setdefault(
                identity,
                {
                    "date": occurrence,
                    "start": start_minute,
                    "end": end_minute,
                    "room_code": room_code,
                    "room_name": actual_name,
                    "building": building,
                    "capacity": capacity,
                    "room_kind": kind,
                    "subject": subject,
                    "type": course_type,
                    "effect_values": set(),
                    "raw_effects": set(),
                    "diplomas": set(),
                    "memos": set(),
                    "source_rows": 0,
                    "period": str(row.get("PERIODE", "")).strip(),
                },
            )
            entry["source_rows"] += 1
            if effect is not None:
                entry["effect_values"].add(effect)
            entry["raw_effects"].add(raw_effect)
            if diploma:
                entry["diplomas"].add(diploma)
            memo = str(row.get("MEMO", "")).strip()
            if memo:
                entry["memos"].add(memo)

    events: list[dict[str, Any]] = []
    for index, item in enumerate(sorted(grouped.values(), key=lambda x: (x["date"], x["start"], x["room_code"], x["subject"])), start=1):
        effects = item.pop("effect_values")
        if len(effects) == 1:
            effect = next(iter(effects))
            effect_issue = ""
        elif not effects:
            effect = None
            effect_issue = "effectif absent ou égal à zéro"
        else:
            effect = None
            effect_issue = "valeurs EFFCALCU contradictoires pour le même créneau/salle"
        item["effect"] = effect
        item["effect_issue"] = effect_issue
        item["diplomas"] = sorted(item["diplomas"], key=str.casefold)
        item["memos"] = sorted(item["memos"])
        item["id"] = f"E{index:07d}"
        item["duration"] = item["end"] - item["start"]
        events.append(item)

    if not events:
        raise SourceError("Aucune séance localisée par CODE_SAL ne reste dans la période/calendrier sélectionnés.")
    return {
        "events": events,
        "active_dates": sorted(schedule_dates),
        "source_rows": raw_rows,
        "source_date_min": date_min,
        "source_date_max": date_max,
        "blank_counts": blank_counts,
        "headers": headers,
        "iso_mismatch": iso_mismatch,
        "day_mismatch": day_mismatch,
        "duration_mismatch": duration_mismatch,
        "at_room_count": at_room_count,
        "at_name_count": at_name_count,
        "at_code_count": at_code_count,
        "period_values": sorted(period_values),
        "period_range_rows": period_range_rows,
        "duration_min": duration_min,
        "duration_max": duration_max,
        "admin_count": admin_count,
        "holiday_count": holiday_count,
        "closure_count": closure_count,
        "weekend_count": weekend_count,
        "excluded_week_count": excluded_week_count,
        "outside_range_count": outside_range_count,
        "missing_room_code_count": missing_room_code_count,
        "unknown_codes": unknown_codes,
        "prefixed_normalized": prefixed_normalized,
        "prefixed_rejected": prefixed_rejected,
        "unrecognized_promotions": sorted(unrecognized_promotions, key=str.casefold),
        "holidays": holidays | closed_dates,
    }


def overlap(a: dict[str, Any], b: dict[str, Any]) -> bool:
    return a["date"] == b["date"] and a["start"] < b["end"] and b["start"] < a["end"]


def clipped_minutes(start: int, end: int) -> int:
    return max(0, min(end, 18 * 60) - max(start, 8 * 60))


def room_events(events: list[dict[str, Any]]) -> dict[tuple[str, dt.date], list[dict[str, Any]]]:
    grouped: dict[tuple[str, dt.date], list[dict[str, Any]]] = defaultdict(list)
    for event in events:
        grouped[(event["room_code"], event["date"])].append(event)
    for items in grouped.values():
        items.sort(key=lambda event: (event["start"], event["end"], event["id"]))
    return grouped


def find_conflicts(events: list[dict[str, Any]], room_codes: set[str] | None = None) -> list[dict[str, Any]]:
    grouped = room_events(events)
    conflicts: list[dict[str, Any]] = []
    for (code, date), items in grouped.items():
        if room_codes is not None and code not in room_codes:
            continue
        for index, first in enumerate(items):
            for second in items[index + 1 :]:
                if second["start"] >= first["end"]:
                    break
                if overlap(first, second):
                    conflicts.append(
                        {
                            "code": code,
                            "room": first["room_name"],
                            "building": first["building"],
                            "date": date.isoformat(),
                            "start1": first["start"],
                            "end1": first["end"],
                            "title1": first["subject"],
                            "type1": first["type"],
                            "diploma1": " | ".join(first["diplomas"]),
                            "start2": second["start"],
                            "end2": second["end"],
                            "title2": second["subject"],
                            "type2": second["type"],
                            "diploma2": " | ".join(second["diplomas"]),
                            "overlap": min(first["end"], second["end"]) - max(first["start"], second["start"]),
                        }
                    )
    return conflicts


def room_free(room_code: str, date: dt.date, start: int, end: int, occupied: dict[tuple[str, dt.date], list[dict[str, Any]]]) -> bool:
    probe = {"date": date, "start": start, "end": end}
    return all(not overlap(probe, existing) for existing in occupied.get((room_code, date), []))


def register_occupied(event: dict[str, Any], occupied: dict[tuple[str, dt.date], list[dict[str, Any]]]) -> None:
    occupied[(event["room_code"], event["date"])].append(event)


def promotion_busy_map(events: list[dict[str, Any]]) -> dict[tuple[str, dt.date], list[tuple[int, int]]]:
    """Créneaux déjà occupés par chaque promotion (NOM_DIP), pour interdire deux séances simultanées."""
    busy: dict[tuple[str, dt.date], list[tuple[int, int]]] = defaultdict(list)
    for event in events:
        for diploma in event["diplomas"]:
            busy[(diploma, event["date"])].append((event["start"], event["end"]))
    return busy


def promotion_free(
    diplomas: list[str], date: dt.date, start: int, end: int, busy: dict[tuple[str, dt.date], list[tuple[int, int]]]
) -> bool:
    return all(
        not (start < b_end and b_start < end)
        for diploma in diplomas
        for b_start, b_end in busy.get((diploma, date), [])
    )


def copy_placement(source: dict[str, Any], room: Room, date: dt.date, start: int, status: str, scenario: str) -> dict[str, Any]:
    placed = dict(source)
    placed.update(
        {
            "id": f"{source['id']}@{scenario}",
            "source_id": source["id"],
            "source_room_code": source["room_code"],
            "source_room_name": source["room_name"],
            "room_code": room.code,
            "room_name": room.name,
            "building": room.building,
            "capacity": room.capacity,
            "room_kind": room.kind,
            "date": date,
            "start": start,
            "end": start + source["duration"],
            "moved": True,
            "status": status,
        }
    )
    return placed


def scenario_result(
    label: str,
    baseline: list[dict[str, Any]],
    sources: list[dict[str, Any]],
    recipients: list[Room],
    active_dates: list[dt.date],
    flexible: bool,
) -> dict[str, Any]:
    occupied = room_events(baseline)
    promo_busy = promotion_busy_map(baseline)
    placed: list[dict[str, Any]] = []
    status_by_id: dict[str, dict[str, Any]] = {}
    room_list = sorted((room for room in recipients if room.capacity is not None and room.capacity > 0), key=lambda r: (r.capacity, r.code))
    ordered_sources = sorted(sources, key=lambda e: (e["date"], e["start"], -e["duration"], e["id"]))

    for source in ordered_sources:
        result = {"status": "", "destination": "", "destination_code": "", "date": source["date"], "start": source["start"], "reason": "", "estimated": bool(source.get("effect_estimated"))}
        if source["effect"] is None:
            result.update(status="effectif_inconnu", reason=source["effect_issue"])
            status_by_id[source["id"]] = result
            continue
        fitting = [room for room in room_list if room.capacity is not None and room.capacity >= source["effect"]]
        if not fitting:
            result.update(status="non_couvert", reason="aucun local d'accueil n'a une capacité suffisante")
            status_by_id[source["id"]] = result
            continue

        if not flexible:
            available = [room for room in fitting if room_free(room.code, source["date"], source["start"], source["end"], occupied)]
            if available:
                chosen = available[0]
                status = "place"
            else:
                chosen = fitting[0]
                status = "place_avec_chevauchement"
            placement = copy_placement(source, chosen, source["date"], source["start"], status, label)
            placed.append(placement)
            register_occupied(placement, occupied)
            result.update(status=status, destination=chosen.name, destination_code=chosen.code)
            status_by_id[source["id"]] = result
            continue

        if source["duration"] > 10 * 60:
            result.update(status="sans_solution", reason="durée supérieure à la plage quotidienne 08:00–18:00")
            status_by_id[source["id"]] = result
            continue

        best: tuple[Any, ...] | None = None
        original_absolute = source["date"].toordinal() * 24 * 60 + source["start"]
        for room in fitting:
            for date in active_dates:
                intervals = sorted(
                    (max(8 * 60, item["start"]), min(18 * 60, item["end"]))
                    for item in occupied.get((room.code, date), [])
                    if item["start"] < 18 * 60 and item["end"] > 8 * 60
                )
                gaps: list[tuple[int, int]] = []
                cursor = 8 * 60
                for left, right in intervals:
                    if left > cursor:
                        gaps.append((cursor, left))
                    cursor = max(cursor, right)
                if cursor < 18 * 60:
                    gaps.append((cursor, 18 * 60))
                for gap_start, gap_end in gaps:
                    if gap_end - gap_start < source["duration"]:
                        continue
                    latest = gap_end - source["duration"]
                    starts = {min(max(source["start"], gap_start), latest), gap_start, latest}
                    for diploma in source["diplomas"]:
                        for b_start, b_end in promo_busy.get((diploma, date), []):
                            starts.update((b_end, b_start - source["duration"]))
                    for chosen_start in sorted(starts):
                        if chosen_start < gap_start or chosen_start > latest:
                            continue
                        if not promotion_free(source["diplomas"], date, chosen_start, chosen_start + source["duration"], promo_busy):
                            continue
                        absolute = date.toordinal() * 24 * 60 + chosen_start
                        score = (abs(absolute - original_absolute), room.capacity, date, chosen_start, room.code)
                        if best is None or score < best[0]:
                            best = (score, room, date, chosen_start)
        if best is None:
            result.update(status="sans_solution", reason="aucun créneau libre (salle, capacité, promotions) trouvé entre 08:00 et 18:00")
        else:
            _, chosen, date, chosen_start = best
            placement = copy_placement(source, chosen, date, chosen_start, "place", label)
            placed.append(placement)
            register_occupied(placement, occupied)
            for diploma in source["diplomas"]:
                promo_busy[(diploma, date)].append((chosen_start, chosen_start + source["duration"]))
            result.update(status="place", destination=chosen.name, destination_code=chosen.code, date=date, start=chosen_start)
        status_by_id[source["id"]] = result

    final_events = [event for event in baseline if event.get("room_code") not in {""}] + placed
    source_minutes = sum(source["duration"] for source in sources)
    accounted_minutes = sum(event["duration"] for event in placed) + sum(
        source["duration"]
        for source in sources
        if status_by_id[source["id"]]["status"] not in {"place", "place_avec_chevauchement"}
    )
    if source_minutes != accounted_minutes:
        raise AnalysisError(f"{label} : conservation des séances échouée ({source_minutes} != {accounted_minutes} minutes).")
    for event in placed:
        if event["effect"] is None or event["capacity"] is None or event["effect"] > event["capacity"]:
            raise AnalysisError(f"{label} : capacité dépassée pour {event['subject']!r} dans {event['room_name']}.")
    return {
        "label": label,
        "events": final_events,
        "placed": placed,
        "statuses": status_by_id,
        "source_minutes": source_minutes,
        "accounted_minutes": accounted_minutes,
        "conflicts": find_conflicts(final_events, {room.code for room in recipients}),
    }


def active_minutes(event: dict[str, Any]) -> int:
    return clipped_minutes(event["start"], event["end"])


def calculate_room_stats(events: list[dict[str, Any]], rooms: dict[str, Room], active_dates: list[dt.date]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    room_map = dict(rooms)
    for event in events:
        if event["room_code"] not in room_map:
            room_map[event["room_code"]] = Room(
                event["room_code"], event["room_name"], event["capacity"], event["building"], event["room_kind"]
            )
    total_minutes: Counter[str] = Counter()
    semester_minutes: Counter[tuple[str, str]] = Counter()
    heat_minutes: Counter[tuple[str, int, int]] = Counter()
    week_minutes: Counter[tuple[str, int, int]] = Counter()
    weekly_days: Counter[tuple[int, int]] = Counter()
    weekday_days: Counter[int] = Counter()
    semester_days: Counter[str] = Counter()
    weekday_semester_days: Counter[tuple[int, str]] = Counter()
    for date in active_dates:
        weekday_days[date.weekday()] += 1
        semester = "S1" if date.month >= 8 else "S2"
        semester_days[semester] += 1
        weekday_semester_days[(date.weekday(), semester)] += 1
        iso = date.isocalendar()
        weekly_days[(iso.year, iso.week)] += 1
    for event in events:
        left = max(8 * 60, event["start"])
        right = min(18 * 60, event["end"])
        if right <= left:
            continue
        minutes = right - left
        code = event["room_code"]
        semester = "S1" if event["date"].month >= 8 else "S2"
        total_minutes[code] += minutes
        semester_minutes[(code, semester)] += minutes
        iso = event["date"].isocalendar()
        week_minutes[(code, iso.year, iso.week)] += minutes
        semester = "S1" if event["date"].month >= 8 else "S2"
        cursor = left
        while cursor < right:
            hour_index = cursor // 60 - 8
            chunk = min(right, (cursor // 60 + 1) * 60) - cursor
            heat_minutes[(code, event["date"].weekday(), hour_index, "ALL")] += chunk
            heat_minutes[(code, event["date"].weekday(), hour_index, semester)] += chunk
            cursor += chunk

    room_rows: list[dict[str, Any]] = []
    for room in sorted(room_map.values(), key=lambda value: (value.building, value.name.casefold(), value.code)):
        denominator = max(1, len(active_dates) * 600)
        s1_denominator = max(1, semester_days["S1"] * 600)
        s2_denominator = max(1, semester_days["S2"] * 600)
        def build_heat(scope: str) -> list[list[float]]:
            output = []
            for weekday in range(5):
                available_days = weekday_days[weekday] if scope == "ALL" else weekday_semester_days[(weekday, scope)]
                day_denominator = max(1, available_days * 60)
                output.append(
                    [round(100 * heat_minutes[(room.code, weekday, hour, scope)] / day_denominator, 6) for hour in range(10)]
                )
            return output

        heat = build_heat("ALL")
        room_rows.append(
            {
                "code": room.code,
                "name": room.name,
                "building": room.building,
                "capacity": room.capacity,
                "kind": room.kind,
                "hours": total_minutes[room.code] / 60,
                "hoursS1": semester_minutes[(room.code, "S1")] / 60,
                "hoursS2": semester_minutes[(room.code, "S2")] / 60,
                "rate": 100 * total_minutes[room.code] / denominator,
                "rateS1": 100 * semester_minutes[(room.code, "S1")] / s1_denominator,
                "rateS2": 100 * semester_minutes[(room.code, "S2")] / s2_denominator,
                "heat": heat,
                "heatS1": build_heat("S1"),
                "heatS2": build_heat("S2"),
            }
        )
    week_rows: list[dict[str, Any]] = []
    for year, week in sorted(weekly_days):
        minutes = sum(value for (code, y, w), value in week_minutes.items() if y == year and w == week)
        week_rows.append(
            {
                "year": year,
                "week": week,
                "days": weekly_days[(year, week)],
                "hours": minutes / 60,
                "rate": 100 * minutes / max(1, len(room_map) * weekly_days[(year, week)] * 600),
                "semester": "S1" if week >= 27 else "S2",
            }
        )
    return room_rows, week_rows


def summarize_scenario(result: dict[str, Any]) -> dict[str, Any]:
    counts = Counter(item["status"] for item in result["statuses"].values())
    hours = Counter()
    for event in result["events"]:
        hours[event["building"]] += active_minutes(event) / 60
    return {
        "label": result["label"],
        "placed": counts["place"] + counts["place_avec_chevauchement"],
        "withConflicts": counts["place_avec_chevauchement"],
        "noSolution": counts["sans_solution"],
        "uncovered": counts["non_couvert"],
        "unknownEffect": counts["effectif_inconnu"],
        "placedHours": sum(event["duration"] for event in result["placed"]) / 60,
        "sourceHours": result["source_minutes"] / 60,
        "accountedHours": result["accounted_minutes"] / 60,
        "conflicts": len(result["conflicts"]),
    }


def independent_manual_check(
    path: Path,
    events: list[dict[str, Any]],
    room_codes: set[str],
    start: dt.date,
    end: dt.date,
    excluded_weeks: set[int],
    holidays: set[dt.date],
) -> list[dict[str, Any]]:
    raw_keys: set[tuple[Any, ...]] = set()
    raw_minutes: Counter[str] = Counter()
    raw_counts: Counter[str] = Counter()
    for row in iter_csv_dicts(path):
        if "__headers__" in row:
            continue
        code = row.get("CODE_SAL", "").strip()
        if code not in room_codes:
            continue
        anchor = dt.datetime.strptime(row["DDEBUT"].strip(), "%d/%m/%Y").date()
        begin_text, finish_text = row["HDEBUT"].strip(), row["HFIN"].strip()
        begin = int(begin_text[:2]) * 60 + int(begin_text[3:])
        finish = int(finish_text[:2]) * 60 + int(finish_text[3:])
        if finish < begin:
            finish += 1440
        if finish - begin >= 480:
            continue
        title = row.get("LIBELLE_MAT", "").strip()
        kind = row.get("TYPE", "").strip()
        for date in dates_from_period(row["PERIODE"].strip(), anchor):
            if date < start or date > end or date in holidays or date.weekday() > 4 or date.isocalendar().week in excluded_weeks:
                continue
            identity = (code, date, begin, finish, title, kind)
            if identity in raw_keys:
                continue
            raw_keys.add(identity)
            raw_minutes[code] += max(0, min(finish, 1080) - max(begin, 480))
            raw_counts[code] += 1

    engine_minutes: Counter[str] = Counter()
    engine_counts: Counter[str] = Counter()
    for event in events:
        if event["room_code"] in room_codes:
            engine_minutes[event["room_code"]] += clipped_minutes(event["start"], event["end"])
            engine_counts[event["room_code"]] += 1
    checks = []
    for code in sorted(room_codes):
        check = {
            "code": code,
            "rawEvents": raw_counts[code],
            "engineEvents": engine_counts[code],
            "rawHours": raw_minutes[code] / 60,
            "engineHours": engine_minutes[code] / 60,
            "differenceMinutes": engine_minutes[code] - raw_minutes[code],
        }
        if check["differenceMinutes"] or check["rawEvents"] != check["engineEvents"]:
            raise AnalysisError(
                f"Rapprochement indépendant échoué pour {code} : {check['rawHours']:.2f} h / "
                f"{check['engineHours']:.2f} h, {check['rawEvents']} / {check['engineEvents']} séances."
            )
        checks.append(check)
    return checks


def html_document(title: str, created_at: str, body: str, css: str | None = None) -> str:
    style = css or ""
    return f"""<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="date" content="{safe_text(created_at)}"><title>{safe_text(title)}</title><style>{style}</style></head><body>{body}</body></html>"""


BASE_CSS = """
:root{color-scheme:light;--ink:#17252a;--muted:#52636b;--line:#d1dadb;--paper:#eef2f1;--surface:#fff;--teal:#17675d;--amber:#b1781b;--red:#a6352f;--radius:7px}
*{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);font:14px/1.48 system-ui,'Segoe UI',sans-serif}main{max-width:1400px;margin:0 auto;padding:24px 22px 52px}header{display:flex;justify-content:space-between;align-items:end;gap:16px;border-bottom:1px solid var(--line);padding:0 0 18px}.eyebrow{color:var(--teal);font:600 11px monospace;text-transform:uppercase}h1{font:600 30px/1.15 Georgia,serif;margin:5px 0}.muted{color:var(--muted)}.stamp{font:12px monospace;text-align:right}.tabs{display:flex;gap:4px;overflow:auto;border-bottom:1px solid var(--line);margin:18px 0 12px}.tab{font:600 13px system-ui;padding:10px 14px;border:0;border-bottom:2px solid transparent;background:none;color:var(--muted);cursor:pointer;white-space:nowrap}.tab[aria-selected=true]{color:var(--teal);border-color:var(--teal)}.filters{display:flex;flex-wrap:wrap;align-items:center;gap:8px;padding:11px;background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);margin-bottom:16px}.filters label{font-size:11px;text-transform:uppercase;color:var(--muted)}input,select{font:inherit;color:var(--ink);background:white;border:1px solid var(--line);border-radius:5px;padding:7px 9px}input[type=search]{flex:1;min-width:180px}.seg{display:flex}.seg button{font:inherit;background:white;border:1px solid var(--line);padding:7px 10px;cursor:pointer}.seg button+button{border-left:0}.seg button[aria-pressed=true]{background:#dcebe7;color:#14594e;font-weight:700}.panel{display:none}.panel.active{display:block}.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin:14px 0}.kpi,section.card{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);padding:15px}.kpi span{display:block;color:var(--muted);font-size:12px}.kpi strong{display:block;font:600 25px/1.2 ui-monospace,monospace;margin-top:7px}.grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}.barrow{display:grid;grid-template-columns:minmax(85px,150px) 1fr 76px;align-items:center;gap:8px;margin:7px 0}.track{height:13px;background:#e6ebeb;border-radius:3px;overflow:hidden}.fill{height:100%;background:var(--teal)}.fill.warn{background:var(--amber)}.fill.crit{background:var(--red)}.tablewrap{overflow:auto;border:1px solid var(--line);border-radius:6px;background:white;max-height:72vh}table{border-collapse:collapse;width:100%;font-size:12px}th,td{padding:7px 8px;text-align:left;white-space:nowrap;border-bottom:1px solid #e5eaea}th{position:sticky;top:0;background:#e8eeee;cursor:pointer;text-transform:uppercase;font-size:10px;color:#415158;z-index:1}tr:hover td{background:#f4f8f7}.num{text-align:right;font-family:ui-monospace,monospace}.pill{display:inline-block;border-radius:99px;padding:2px 7px;background:#e2efeb;color:#14594e;font:600 11px monospace}.pill.warn{background:#fff0d6;color:#805311}.pill.crit{background:#f9e2df;color:#922d27}.heat{display:inline-block;min-width:38px;text-align:center;padding:4px;border-radius:3px;font:11px monospace}.note{padding:12px 14px;border-left:4px solid var(--teal);background:#e0ece9;margin:10px 0}.note.warn{border-color:var(--amber);background:#fff2dc}.note.crit{border-color:var(--red);background:#fae8e5}.bars{display:flex;align-items:end;gap:4px;height:175px;overflow:auto;padding:8px}.weekbar{flex:0 0 27px;min-height:2px;background:var(--teal);position:relative}.weekbar.s2{background:var(--amber)}.weekbar span{position:absolute;bottom:-22px;font:9px monospace}.hidden{display:none!important}a{color:var(--teal)}footer{color:var(--muted);border-top:1px solid var(--line);margin-top:22px;padding-top:12px;font-size:11px}
@media(max-width:800px){main{padding:16px 12px 36px}header{align-items:start;flex-direction:column}.stamp{text-align:left}.grid2{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
"""


def render_dashboard(
    title: str,
    scenario_label: str,
    created_at: str,
    room_rows: list[dict[str, Any]],
    week_rows: list[dict[str, Any]],
    conflicts: list[dict[str, Any]],
    scenario_summary: dict[str, Any],
    active_days: int,
    created_stamp: str,
) -> str:
    payload = json.dumps(
        {"rooms": room_rows, "weeks": week_rows, "conflicts": conflicts, "summary": scenario_summary, "activeDays": active_days},
        ensure_ascii=False,
        separators=(",", ":"),
    ).replace("</", "<\\/")
    body = f"""<main><header><div><div class="eyebrow">UPPA · Direction du Patrimoine · SSH</div><h1>{safe_text(title)}</h1><div class="muted">{safe_text(scenario_label)} · {len(room_rows)} locaux · {active_days} jours pédagogiques observés</div></div><div class="stamp">Créé le<br>{safe_text(created_at)}</div></header>
<nav class="tabs" role="tablist"><button class="tab" data-tab="overview" aria-selected="true">Vue d'ensemble</button><button class="tab" data-tab="grid">Grille horaire</button><button class="tab" data-tab="weekly">Analyse hebdomadaire</button><button class="tab" data-tab="exec">Synthèse exécutive</button><button class="tab" data-tab="conflicts">Chevauchements</button></nav>
<div class="filters"><label for="building">Bâtiment</label><select id="building"><option value="">Tous</option></select><input id="search" type="search" placeholder="Rechercher une salle…"><label for="rate">Taux</label><select id="rate"><option value="">Tous</option><option value="low">&lt; 10 %</option><option value="mid">10–30 %</option><option value="high">≥ 30 %</option><option value="over">&gt; 100 %</option></select><div class="seg"><button data-sem="" aria-pressed="true">Année</button><button data-sem="S1" aria-pressed="false">S1</button><button data-sem="S2" aria-pressed="false">S2</button></div></div>
<section class="panel active" id="panel-overview"><div class="kpis" id="kpis"></div><div class="grid2"><section class="card"><h2>Occupation moyenne par bâtiment</h2><div id="buildings"></div></section><section class="card"><h2>Locaux les plus sollicités</h2><div id="top"></div></section></div><section class="card"><h2>Détail des locaux</h2><div class="tablewrap"><table><thead><tr><th data-key="name">Local</th><th data-key="building">Bâtiment</th><th data-key="capacity" class="num">Places</th><th data-key="hours" class="num">Heures</th><th data-key="rate" class="num">Taux</th></tr></thead><tbody id="room-table"></tbody></table></div></section></section>
<section class="panel" id="panel-grid"><section class="card"><h2>Grille heure × salle</h2><p class="muted">Part de la plage 08:00–18:00 occupée, moyennée sur les jours pédagogiques observés. Les chevauchements peuvent dépasser 100 %.</p><div class="tablewrap"><table><thead id="heat-head"></thead><tbody id="heat-table"></tbody></table></div></section></section>
<section class="panel" id="panel-weekly"><section class="card"><h2>Taux d'occupation hebdomadaire</h2><div class="bars" id="week-bars"></div><div style="height:28px"></div><div class="tablewrap"><table><thead><tr><th>Semaine</th><th class="num">Heures</th><th class="num">Taux</th></tr></thead><tbody id="week-table"></tbody></table></div></section></section>
<section class="panel" id="panel-exec"><div id="exec-summary"></div><section class="card"><h2>Hypothèses de lecture</h2><ul><li>Les heures utilisent les intervalles HDEBUT/HFIN ; la capacité de chaque salle est issue d'EXP_SALLE.</li><li>Les jours fériés nationaux français sont retranchés. Les autres jours sans cours sont déterminés par les dates pédagogiques réellement présentes dans l'export ; semaines exclues sur option.</li><li>EFFCALCU est utilisé tel qu'exporté. Les valeurs nulles/0 ou contradictoires ne sont pas remplacées par une estimation.</li><li>Le lissage est un test de disponibilité théorique, pas un planning opérationnel. Les plateaux d'examen ne sont pas modélisés.</li></ul></section></section>
<section class="panel" id="panel-conflicts"><section class="card"><h2>Chevauchements détectés</h2><div class="muted" id="conflict-count"></div><div class="tablewrap"><table><thead><tr><th>Local</th><th>Bât.</th><th>Date</th><th>Créneau 1</th><th>Séance 1</th><th>Créneau 2</th><th>Séance 2</th><th class="num">Min.</th></tr></thead><tbody id="conflict-table"></tbody></table></div></section></section>
<footer>Créé le {safe_text(created_at)} · Données locales intégrées à cette page ; aucun serveur ni ressource externe requis. Exports : 06 surdimensionnement et 07 détail séance par séance, horodatés {safe_text(created_stamp)}.</footer>
<script type="application/json" id="payload">{payload}</script><script>
(()=>{{const p=JSON.parse(document.getElementById('payload').textContent),$=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)],fmt=(n,d=1)=>Number(n||0).toLocaleString('fr-FR',{{minimumFractionDigits:d,maximumFractionDigits:d}}),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c]));let sem='',sortKey='rate',sortDir=-1;
const rate=r=>sem==='S1'?r.rateS1:sem==='S2'?r.rateS2:r.rate,hrs=r=>sem==='S1'?r.hoursS1:sem==='S2'?r.hoursS2:r.hours;
const bld=$('#building');[...new Set(p.rooms.map(r=>r.building).filter(Boolean))].sort().forEach(v=>bld.add(new Option(v,v)));
function visible(){{let q=$('#search').value.trim().toLowerCase(),building=bld.value,band=$('#rate').value;return p.rooms.filter(r=>{{let t=rate(r);return(!building||r.building===building)&&(!q||(r.name+' '+r.building+' '+r.code).toLowerCase().includes(q))&&(!band||(band==='low'&&t<10)||(band==='mid'&&t>=10&&t<30)||(band==='high'&&t>=30&&t<=100)||(band==='over'&&t>100));}})}}
function render(){{let rows=visible(),sumH=rows.reduce((a,r)=>a+hrs(r),0),den=rows.length*p.activeDays*10,avg=den?100*sumH/den:0,over=rows.filter(r=>rate(r)>100).length;
$('#kpis').innerHTML=[["Locaux filtrés",rows.length],["Taux pondéré",fmt(avg,2)+' %'],["Heures-salles",fmt(sumH,1)+' h'],["Locaux au-dessus de 100 %",over]].map(x=>'<div class="kpi"><span>'+x[0]+'</span><strong>'+x[1]+'</strong></div>').join('');
let gb={{}};rows.forEach(r=>{{let x=gb[r.building||'—']||(gb[r.building||'—']={{h:0,n:0}});x.h+=hrs(r);x.n++;}});$('#buildings').innerHTML=Object.entries(gb).map(([k,x])=>{{let v=x.n*p.activeDays*10?100*x.h/(x.n*p.activeDays*10):0;return '<div class="barrow"><span>'+esc(k)+'</span><span class="track"><span class="fill '+(v>100?'crit':'')+'" style="display:block;width:'+Math.min(100,v)+'%"></span></span><b>'+fmt(v,1)+'%</b></div>';}}).join('')||'<p>Aucune donnée.</p>';
let top=rows.slice().sort((a,b)=>rate(b)-rate(a)).slice(0,8);$('#top').innerHTML=top.map(r=>'<div class="barrow"><span title="'+esc(r.name)+'">'+esc(r.name)+'</span><span class="track"><span class="fill '+(rate(r)>100?'crit':'')+'" style="display:block;width:'+Math.min(100,rate(r))+'%"></span></span><b>'+fmt(rate(r),1)+'%</b></div>').join('')||'<p>Aucune donnée.</p>';
let sorted=rows.slice().sort((a,b)=>{{let x=a[sortKey],y=b[sortKey];if(sortKey==='rate'){{x=rate(a);y=rate(b)}}if(sortKey==='hours'){{x=hrs(a);y=hrs(b)}}return(typeof x==='string'?x.localeCompare(y,'fr'):(x??-1)-(y??-1))*sortDir;}});$('#room-table').innerHTML=sorted.map(r=>'<tr><td>'+esc(r.name)+'</td><td>'+esc(r.building||'—')+'</td><td class="num">'+(r.capacity??'—')+'</td><td class="num">'+fmt(hrs(r))+'</td><td class="num"><span class="pill '+(rate(r)>100?'crit':rate(r)>=80?'warn':'')+'">'+fmt(rate(r),2)+'%</span></td></tr>').join('');
let names=['Lun.','Mar.','Mer.','Jeu.','Ven.'];$('#heat-head').innerHTML='<tr><th>Local</th><th>Bât.</th>'+Array.from({{length:10}},(_,i)=>'<th class="num">'+(i+8)+'–'+(i+9)+'</th>').join('')+'</tr>';
let heatKey=sem==='S1'?'heatS1':sem==='S2'?'heatS2':'heat';$('#heat-table').innerHTML=rows.map(r=>'<tr><td>'+esc(r.name)+'</td><td>'+esc(r.building||'—')+'</td>'+Array.from({{length:10}},(_,i)=>{{let v=r[heatKey].reduce((a,d)=>a+(d[i]||0),0)/5,alpha=Math.min(.85,v/110);return '<td><span class="heat" title="moyenne lun–ven '+fmt(v,1)+'%" style="background:rgba(23,103,93,'+alpha+')">'+fmt(v,0)+'%</span></td>';}}).join('')+'</tr>').join('');
let weekMap={{}};rows.forEach(r=>r.weekly.forEach(w=>{{if(sem&&w.semester!==sem)return;let key=w.year+'-'+w.week,x=weekMap[key]||(weekMap[key]={{year:w.year,week:w.week,days:w.days,semester:w.semester,hours:0}});x.hours+=w.hours;}}));let weeks=Object.values(weekMap).map(w=>({{...w,rate:100*w.hours/Math.max(1,rows.length*w.days*10)}})).sort((a,b)=>a.year-b.year||a.week-b.week),max=Math.max(1,...weeks.map(w=>w.rate));$('#week-bars').innerHTML=weeks.map(w=>'<div class="weekbar '+(w.semester==='S2'?'s2':'')+'" title="'+w.year+' S'+w.week+' · '+fmt(w.rate,1)+'%" style="height:'+Math.max(2,Math.min(100,100*w.rate/max))+'%"><span>S'+w.week+'</span></div>').join('');$('#week-table').innerHTML=weeks.map(w=>'<tr><td>'+w.year+' · S'+w.week+'</td><td class="num">'+fmt(w.hours)+'</td><td class="num">'+fmt(w.rate,2)+'%</td></tr>').join('');
let s=p.summary,filteredHours=rows.reduce((a,r)=>a+hrs(r),0),filteredRate=100*filteredHours/Math.max(1,rows.length*p.activeDays*10);$('#exec-summary').innerHTML='<section class="card"><h2>'+esc(s.label)+'</h2><p>Filtre courant : '+rows.length+' locaux, '+fmt(filteredHours,1)+' h, taux pondéré '+fmt(filteredRate,2)+'%.</p><p>Bilan global du scénario : '+s.placed+' séances placées, dont '+s.withConflicts+' avec chevauchement brut ; '+s.uncovered+' non couvertes par capacité, '+s.noSolution+' sans solution au lissage, '+s.unknownEffect+' sans effectif déterminé.</p><p>Heures sources suivies : '+fmt(s.sourceHours,2)+' h ; heures effectivement placées : '+fmt(s.placedHours,2)+' h. Les heures non placées restent listées dans le détail.</p><p>Chevauchements de locaux : '+s.conflicts+'.</p></section>';
let roomByCode=Object.fromEntries(p.rooms.map(r=>[r.code,r])),q=$('#search').value.trim().toLowerCase();let cs=p.conflicts.filter(c=>{{let room=roomByCode[c.code],t=room?rate(room):null,month=Number(c.date.slice(5,7)),confSem=month>=8?'S1':'S2',band=$('#rate').value;return(!sem||confSem===sem)&&(!bld.value||c.building===bld.value)&&(!q||((c.room+' '+c.building+' '+c.title1+' '+c.title2).toLowerCase().includes(q)))&&(!band||(band==='low'&&t<10)||(band==='mid'&&t>=10&&t<30)||(band==='high'&&t>=30&&t<=100)||(band==='over'&&t>100));}});$('#conflict-count').textContent=cs.length+' chevauchements sur le périmètre filtré';$('#conflict-table').innerHTML=cs.slice(0,5000).map(c=>'<tr><td>'+esc(c.room)+'</td><td>'+esc(c.building)+'</td><td>'+esc(c.date)+'</td><td>'+esc(c.slot1)+'</td><td>'+esc(c.title1)+'</td><td>'+esc(c.slot2)+'</td><td>'+esc(c.title2)+'</td><td class="num">'+c.overlap+'</td></tr>').join('');}}
$$('.tab').forEach(b=>b.addEventListener('click',()=>{{$$('.tab').forEach(x=>x.setAttribute('aria-selected',String(x===b)));$$('.panel').forEach(x=>x.classList.toggle('active',x.id==='panel-'+b.dataset.tab));}}));
$$('.seg button').forEach(b=>b.addEventListener('click',()=>{{sem=b.dataset.sem;$$('.seg button').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));render();}}));['#building','#search','#rate'].forEach(s=>$(s).addEventListener('input',render));$$('th[data-key]').forEach(th=>th.addEventListener('click',()=>{{sortDir=sortKey===th.dataset.key?-sortDir:1;sortKey=th.dataset.key;render();}}));render();}})();
</script></main>"""
    return html_document(title, created_at, body, BASE_CSS)


# Correctif de lisibilité du gabarit : texte clair sur fond sombre et texte sombre sur fond clair (contraste >= 4,5:1).
CONTRAST_FIX_CSS = """<style id="contrast-fix">
:root{--ink-faint:#5a6671;--warning:#8f5c0c;--brass:#86601d;--heat-4:#25757d}
.heatcell.heat-4,.heatcell.heat-5,.heatcell.heat-x{color:#fff}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--ink-faint:#92a0aa;--warning:#dea748;--brass:#d6a555;--heat-4:#2f8890}
:root:not([data-theme="light"]) .heatcell.heat-4,:root:not([data-theme="light"]) .heatcell.heat-5,:root:not([data-theme="light"]) .heatcell.heat-x,:root:not([data-theme="light"]) .pstep .circle{color:#0b1115}}
:root[data-theme="dark"]{--ink-faint:#92a0aa;--warning:#dea748;--brass:#d6a555;--heat-4:#2f8890}
:root[data-theme="dark"] .heatcell.heat-4,:root[data-theme="dark"] .heatcell.heat-5,:root[data-theme="dark"] .heatcell.heat-x,:root[data-theme="dark"] .pstep .circle{color:#0b1115}
</style>
"""


def legacy_template_dashboard(
    template_path: Path,
    scenario: str,
    created_at: str,
    room_rows: list[dict[str, Any]],
    week_rows: list[dict[str, Any]],
    events: list[dict[str, Any]],
    conflicts: list[dict[str, Any]],
    active_dates: list[dt.date],
    building_sites: dict[str, dict[str, str]] | None = None,
    exam_periods: list[tuple[dt.date, dt.date]] | None = None,
) -> str:
    building_sites = building_sites or {}
    exam_periods = exam_periods or []
    template = template_path.read_text(encoding="utf-8")
    template = re.sub(r'<link rel="preconnect" href="https://fonts\.gstatic\.com" crossorigin>\s*', "", template)
    # Placé tout à la fin du document pour passer après les styles du gabarit (qui peuvent aussi être déclarés après </head>).
    body_end = template.rfind("</body>")
    template = template[:body_end] + CONTRAST_FIX_CSS + template[body_end:] if body_end >= 0 else template + CONTRAST_FIX_CSS
    template = re.sub(r'<link rel="stylesheet" href="https://fonts\.googleapis\.com/[^>]+>\s*', "", template)
    template = template.replace(
        "document.getElementById('exec-creneaux').textContent = semestre==='S1' ? '690' : semestre==='S2' ? '790' : '1 480';",
        "document.getElementById('exec-creneaux').textContent = semestre==='S1' ? String({{JOURS_S1}}*10) : semestre==='S2' ? String({{JOURS_S2}}*10) : String({{JOURS_TOTAL}}*10);",
    )

    exam_supported = 'data-sem="S2"' in template  # gabarits de test sans sélecteur de semestre : rien à compléter

    def patch(old: str, new: str) -> None:
        nonlocal template
        if not exam_supported:
            return
        if old not in template:
            raise AnalysisError(f"Gabarit modifié : impossible d'ajouter les périodes d'examen (motif introuvable : {old[:60]!r}).")
        template = template.replace(old, new, 1)

    patch(
        '<button type="button" data-sem="S2">2nd sem.</button>',
        '<button type="button" data-sem="S2">2nd sem.</button>\n      '
        '<button type="button" data-sem="EX" title="Périodes d\'examen (début janvier, fin mai–juin) — modifiables avec --examens">Examens</button>\n      '
        '<button type="button" data-sem="HX" title="Tous les jours pédagogiques hors périodes d\'examen">Hors examens</button>',
    )
    patch(
        "const heatmapSalleS2 = {{HEATMAPSALLES2}}\n;",
        "const heatmapSalleS2 = {{HEATMAPSALLES2}}\n;\n  const heatmapSalleEX = {{HEATMAPSALLEEX}};\n  const heatmapSalleHX = {{HEATMAPSALLEHX}};\n"
        "  const EXAM_WEEKS = new Set({{EXAMWEEKS}});\n  const EXAM_DATES = new Set({{EXAMDATES}});\n  const EXAM_LABEL = {{EXAMLABEL_JSON}};\n",
    )
    patch(
        "heuresS2: r.heuresS2 ?? 0, tauxS2: r.tauxS2\n",
        "heuresS2: r.heuresS2 ?? 0, tauxS2: r.tauxS2,\n    heuresEX: r.heuresEX ?? 0, tauxEX: r.tauxEX, heuresHX: r.heuresHX ?? 0, tauxHX: r.tauxHX\n",
    )
    patch(
        "function eTaux(d){ return semestre==='S1' ? d.tauxS1 : semestre==='S2' ? d.tauxS2 : d.taux; }",
        "function eTaux(d){ return semestre==='S1' ? d.tauxS1 : semestre==='S2' ? d.tauxS2 : semestre==='EX' ? d.tauxEX : semestre==='HX' ? d.tauxHX : d.taux; }",
    )
    patch(
        "function eHeures(d){ return semestre==='S1' ? (d.heuresS1||0) : semestre==='S2' ? (d.heuresS2||0) : (d.heures||0); }",
        "function eHeures(d){ return semestre==='S1' ? (d.heuresS1||0) : semestre==='S2' ? (d.heuresS2||0) : semestre==='EX' ? (d.heuresEX||0) : semestre==='HX' ? (d.heuresHX||0) : (d.heures||0); }",
    )
    patch(
        "function semLabel(){ return semestre==='S1' ?",
        "function semLabel(){ return semestre==='EX' ? EXAM_LABEL : semestre==='HX' ? 'les jours hors périodes d\\'examen' : semestre==='S1' ?",
    )
    patch(
        "    if (semestre === 'S2') weeks = weeks.filter(w => w.annee === {{ANNEE_S2}});\n",
        "    if (semestre === 'S2') weeks = weeks.filter(w => w.annee === {{ANNEE_S2}});\n"
        "    if (semestre === 'EX') weeks = weeks.filter(w => EXAM_WEEKS.has(w.label));\n"
        "    if (semestre === 'HX') weeks = weeks.filter(w => !EXAM_WEEKS.has(w.label));\n",
    )
    patch(
        "const heatmapSource = semestre === 'S1' ? heatmapSalleS1 : semestre === 'S2' ? heatmapSalleS2 : heatmapSalle;",
        "const heatmapSource = semestre === 'S1' ? heatmapSalleS1 : semestre === 'S2' ? heatmapSalleS2 : semestre === 'EX' ? heatmapSalleEX : semestre === 'HX' ? heatmapSalleHX : heatmapSalle;",
    )
    patch(
        "semestre==='S1' ? '{{JOURS_S1}}' : semestre==='S2' ? '{{JOURS_S2}}' : '{{JOURS_TOTAL}}'",
        "semestre==='S1' ? '{{JOURS_S1}}' : semestre==='S2' ? '{{JOURS_S2}}' : semestre==='EX' ? '{{JOURS_EX}}' : semestre==='HX' ? '{{JOURS_HX}}' : '{{JOURS_TOTAL}}'",
    )
    patch(
        "semestre==='S1' ? String({{JOURS_S1}}*10) : semestre==='S2' ? String({{JOURS_S2}}*10) : String({{JOURS_TOTAL}}*10)",
        "semestre==='S1' ? String({{JOURS_S1}}*10) : semestre==='S2' ? String({{JOURS_S2}}*10) : semestre==='EX' ? String({{JOURS_EX}}*10) : semestre==='HX' ? String({{JOURS_HX}}*10) : String({{JOURS_TOTAL}}*10)",
    )
    patch(
        "    if (semestre === 'S2' && c.annee !== {{ANNEE_S2}}) return false;\n",
        "    if (semestre === 'S2' && c.annee !== {{ANNEE_S2}}) return false;\n"
        "    if (semestre === 'EX' && !EXAM_DATES.has(c.date)) return false;\n"
        "    if (semestre === 'HX' && EXAM_DATES.has(c.date)) return false;\n",
    )

    counts: Counter[str] = Counter(event["room_code"] for event in events)
    key_counts = Counter((room["name"], room["building"]) for room in room_rows)
    labels: dict[str, str] = {}
    raw_data: list[dict[str, Any]] = []
    heat_all: dict[str, Any] = {}
    heat_s1: dict[str, Any] = {}
    heat_s2: dict[str, Any] = {}
    weekday_names = ("lundi", "mardi", "mercredi", "jeudi", "vendredi")

    def heat_object(values: list[list[float]]) -> dict[str, Any]:
        return {
            day: {f"{hour + 8}h-{hour + 9}h": values[weekday][hour] for hour in range(10)}
            for weekday, day in enumerate(weekday_names)
        }

    exam_dates = [date for date in active_dates if is_exam_date(date, exam_periods)]
    exam_set = set(exam_dates)
    other_dates = [date for date in active_dates if date not in exam_set]
    room_catalog = {
        room["code"]: Room(room["code"], room["name"], room["capacity"], room["building"], room["kind"]) for room in room_rows
    }

    def period_stats(dates: list[dt.date]) -> dict[str, dict[str, Any]]:
        chosen = set(dates)
        rows, _ = calculate_room_stats([event for event in events if event["date"] in chosen], room_catalog, dates)
        return {row["code"]: row for row in rows}

    exam_stats, other_stats = period_stats(exam_dates), period_stats(other_dates)
    heat_ex: dict[str, Any] = {}
    heat_hx: dict[str, Any] = {}
    for room in room_rows:
        key = f"{room['name']}|{room['building']}"
        label = room["name"]
        if key_counts[(room["name"], room["building"])] > 1:
            label = f"{label} [{room['code']}]"
            key = f"{label}|{room['building']}"
        labels[room["code"]] = key
        raw_data.append(
            {
                "salle": label,
                "batiment": room["building"],
                "ville": building_sites.get(room["building"], {}).get("ville", ""),
                "site": building_sites.get(room["building"], {}).get("site", ""),
                "etage": "",
                "capa": room["capacity"],
                "surface": None,
                "ratio": None,
                "heures": room["hours"],
                "nbCours": counts[room["code"]],
                "taux": room["rate"],
                "heuresS1": room["hoursS1"],
                "tauxS1": room["rateS1"],
                "heuresS2": room["hoursS2"],
                "tauxS2": room["rateS2"],
                "heuresEX": exam_stats[room["code"]]["hours"],
                "tauxEX": exam_stats[room["code"]]["rate"],
                "heuresHX": other_stats[room["code"]]["hours"],
                "tauxHX": other_stats[room["code"]]["rate"],
            }
        )
        heat_all[key] = heat_object(room["heat"])
        heat_s1[key] = heat_object(room["heatS1"])
        heat_s2[key] = heat_object(room["heatS2"])
        heat_ex[key] = heat_object(exam_stats[room["code"]]["heat"])
        heat_hx[key] = heat_object(other_stats[room["code"]]["heat"])

    heat_by_building: dict[str, list[list[float]]] = {}
    for source_room, room in zip(room_rows, raw_data):
        building = room["batiment"] or "Autres"
        if building not in heat_by_building:
            heat_by_building[building] = [[0.0 for _ in range(10)] for _ in range(5)]
        key = labels[source_room["code"]]
        heat = heat_all[key]
        for weekday, day in enumerate(weekday_names):
            for hour in range(10):
                heat_by_building[building][weekday][hour] += heat[day][f"{hour + 8}h-{hour + 9}h"]
    heatmap_building = {}
    for building, matrix in heat_by_building.items():
        room_count = sum(1 for room in raw_data if (room["batiment"] or "Autres") == building)
        averaged = [[value / max(1, room_count) for value in row] for row in matrix]
        heatmap_building[building] = heat_object(averaged)

    week_days = Counter((date.isocalendar().year, date.isocalendar().week) for date in active_dates)
    weekly_all: list[dict[str, Any]] = []
    for week in week_rows:
        year, number = week["year"], week["week"]
        try:
            monday = dt.date.fromisocalendar(year, number, 1)
        except ValueError:
            continue
        weekly_all.append(
            {
                "label": f"{year}-S{number:02d}",
                "annee": year,
                "semaine": number,
                "nbJoursOuvresSemaine": week_days[(year, number)],
                "lundi": monday.strftime("%d/%m/%Y"),
            }
        )
    weekly_salle: dict[str, list[dict[str, Any]]] = {}
    for room in room_rows:
        key = labels[room["code"]]
        weekly_salle[key] = [
            {
                "label": f"{week['year']}-S{week['week']:02d}",
                "annee": week["year"],
                "semaine": week["week"],
                "heures": week["hours"],
            }
            for week in room["weekly"]
        ]

    raw_conflicts = []
    french_days = ("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")
    for conflict in conflicts:
        date = dt.date.fromisoformat(conflict["date"])
        raw_conflicts.append(
            {
                "salle": conflict["room"],
                "batiment": conflict["building"],
                "date": date.strftime("%d/%m/%Y"),
                "jour": french_days[date.weekday()],
                "hdebut1": f"{conflict['start1']//60:02d}:{conflict['start1']%60:02d}",
                "hfin1": f"{conflict['end1']//60:02d}:{conflict['end1']%60:02d}",
                "matiere1": conflict["title1"],
                "diplome1": conflict["diploma1"],
                "hdebut2": f"{conflict['start2']//60:02d}:{conflict['start2']%60:02d}",
                "hfin2": f"{conflict['end2']//60:02d}:{conflict['end2']%60:02d}",
                "matiere2": conflict["title2"],
                "diplome2": conflict["diploma2"],
                "overlapMin": conflict["overlap"],
            }
        )

    first_date, last_date = min(active_dates), max(active_dates)
    sem1_days = sum(date.month >= 8 for date in active_dates)
    sem2_days = len(active_dates) - sem1_days
    values = {
        "DATE_CREATION_ISO": safe_text(created_at),
        "DATE_CREATION": safe_text(dt.datetime.fromisoformat(created_at).strftime("%d/%m/%Y à %H:%M:%S")),
        "DATE_EXTRACTION": "non renseignée dans le CSV fourni",
        "SCENARIO": safe_text(scenario),
        "NB_SALLES": str(len(raw_data)),
        "PERIODE_LABEL": safe_text(f"{first_date.strftime('%d/%m/%Y')} – {last_date.strftime('%d/%m/%Y')}"),
        "NOTE_CHEVAUCHEMENTS": safe_text("Les conflits sont conservés et détaillés séance par séance ; une occupation supérieure à 100 % signale des chevauchements."),
        "NOTE_REFERENCE": safe_text(" Les surfaces ne figurent pas dans EXP_SALLE et restent non renseignées."),
        "ANNEE_S1": str(first_date.year),
        "ANNEE_S2": str(last_date.year),
        "JOURS_S1": str(sem1_days),
        "JOURS_S2": str(sem2_days),
        "JOURS_TOTAL": str(len(active_dates)),
        "HEURES_BASE": "10",
        "RAWDATA": json.dumps(raw_data, ensure_ascii=False, separators=(",", ":")),
        "HEATMAPBAT": json.dumps(heatmap_building, ensure_ascii=False, separators=(",", ":")),
        "HEATMAPSALLE": json.dumps(heat_all, ensure_ascii=False, separators=(",", ":")),
        "HEATMAPSALLES1": json.dumps(heat_s1, ensure_ascii=False, separators=(",", ":")),
        "HEATMAPSALLES2": json.dumps(heat_s2, ensure_ascii=False, separators=(",", ":")),
        "HEATMAPSALLEEX": json.dumps(heat_ex, ensure_ascii=False, separators=(",", ":")),
        "HEATMAPSALLEHX": json.dumps(heat_hx, ensure_ascii=False, separators=(",", ":")),
        "EXAMWEEKS": json.dumps(
            sorted({f"{date.isocalendar().year}-S{date.isocalendar().week:02d}" for date in exam_dates}), ensure_ascii=False
        ),
        "EXAMDATES": json.dumps(sorted(date.strftime("%d/%m/%Y") for date in exam_dates), ensure_ascii=False),
        "EXAMLABEL_JSON": json.dumps(
            "les périodes d'examen (" + " ; ".join(f"{a.strftime('%d/%m/%Y')} – {b.strftime('%d/%m/%Y')}" for a, b in exam_periods) + ")"
            if exam_periods else "les périodes d'examen (aucune définie)",
            ensure_ascii=False,
        ),
        "JOURS_EX": str(len(exam_dates)),
        "JOURS_HX": str(len(other_dates)),
        "WEEKLYDATA": json.dumps({"TOUS": weekly_all}, ensure_ascii=False, separators=(",", ":")),
        "RAWCONFLICTS": json.dumps(raw_conflicts, ensure_ascii=False, separators=(",", ":")),
        "WEEKLYSALLEDATA": json.dumps(weekly_salle, ensure_ascii=False, separators=(",", ":")),
    }
    for key in ("RAWDATA", "HEATMAPBAT", "HEATMAPSALLE", "HEATMAPSALLES1", "HEATMAPSALLES2", "HEATMAPSALLEEX", "HEATMAPSALLEHX", "EXAMWEEKS", "EXAMDATES", "WEEKLYDATA", "RAWCONFLICTS", "WEEKLYSALLEDATA"):
        values[key] = values[key].replace("</", "<\\/")
    placeholders = set(re.findall(r"\{\{([A-Z0-9_]+)\}\}", template))
    missing_values = placeholders - values.keys()
    if missing_values:
        raise AnalysisError(f"Placeholders inconnus dans le gabarit : {', '.join(sorted(missing_values))}.")
    template = re.sub(
        r"\{\{([A-Z0-9_]+)\}\}",
        lambda match: values[match.group(1)],
        template,
    )
    leftovers = re.findall(r"\{\{([A-Z0-9_]+)\}\}", template)
    if leftovers:
        raise AnalysisError(f"Placeholders non remplacés dans le gabarit : {', '.join(sorted(set(leftovers)))}.")
    return template


CASE_EXPLANATIONS: dict[str, dict[str, str]] = {
    "amphis": {
        "title": "Les amphithéâtres",
        "what": "La liste des amphis, avec le pourcentage de temps où chacun est utilisé aujourd'hui (Référence) et dans chaque simulation (H1 à H3b).",
        "how": "Choisissez une ville, un site ou un bâtiment avec les listes. Cochez des amphis pour voir leur moyenne. Cliquez sur le titre d'une colonne pour trier.",
    },
    "sans_effectif": {
        "title": "Les cours dont on ne connaît pas le nombre d'étudiants",
        "what": "Ces cours des amphis de Lettres n'ont pas pu être déplacés, car le planning ne dit pas combien d'étudiants y assistent.",
        "how": "La page explique pourquoi, donne la liste complète (triable, avec recherche) et le moyen de les placer : renseigner leur effectif dans le fichier CSV.",
    },
    "reference": {
        "title": "Référence : la situation actuelle",
        "what": "Comment les salles sont utilisées aujourd'hui, sans rien changer.",
        "how": "C'est le point de départ pour comparer avec H1, H2, H3a et H3b. Un taux au-dessus de 100 % veut dire que deux cours étaient prévus en même temps dans la même salle.",
    },
    "H1": {
        "title": "H1 : les cours des amphis de Lettres vont dans les amphis de Droit, aux mêmes horaires",
        "what": "On supprime les 3 amphis de Lettres. Chaque cours va dans un des 5 amphis de Droit, le même jour à la même heure : on prend le plus petit amphi qui est libre et assez grand.",
        "how": "Si aucun amphi n'est libre à cette heure, le cours est quand même placé et deux cours se retrouvent en même temps dans la même salle. C'est le cas le plus simple, donc le plus risqué.",
    },
    "H2": {
        "title": "H2 : amphis de Droit, avec des horaires qui peuvent changer",
        "what": "Comme H1, mais un cours peut être décalé à une autre heure ou un autre jour (le plus proche possible de l'horaire d'origine).",
        "how": "Un cours n'est décalé que si la salle est libre et assez grande, entre 8 h et 18 h, et si sa promotion n'a pas déjà un autre cours à ce moment. C'est un test, pas un emploi du temps officiel.",
    },
    "H3a": {
        "title": "H3a : amphis de Droit + grandes salles, aux mêmes horaires",
        "what": "Comme H1, mais les cours peuvent aussi aller dans les grandes salles de cours de Lettres et de Droit (de 50 à 100 places).",
        "how": "Il y a plus de salles possibles, donc moins de cours en même temps dans la même salle qu'en H1.",
    },
    "H3b": {
        "title": "H3b : amphis de Droit + grandes salles, avec des horaires qui peuvent changer",
        "what": "Comme H2, mais avec les grandes salles de cours en plus des amphis.",
        "how": "C'est le cas le plus favorable. Cela reste un test théorique.",
    },
    "surdimensionnement": {
        "title": "Salles trop grandes pour leur cours",
        "what": "Les cours qui ont lieu dans une salle bien trop grande : au moins deux fois plus de places que d'étudiants, et au moins 20 places vides.",
        "how": "Pour chaque cas, on propose une salle plus petite, libre à toutes les dates du cours. Si aucune n'est libre sur toute la série, rien n'est proposé.",
    },
}

STATUS_EXPLANATIONS: dict[str, str] = {
    "place": "Le cours a trouvé une salle libre et assez grande.",
    "place_avec_chevauchement": "H1 et H3a : aucune salle n'est libre à cette heure. Le cours est placé quand même, donc deux cours ont lieu en même temps.",
    "sans_solution": "H2 et H3b : aucune heure libre trouvée. Le cours n'est pas placé.",
    "non_couvert": "Aucune salle n'est assez grande pour le nombre d'étudiants. Le cours n'est pas placé.",
    "effectif_inconnu": "On ne connaît pas le nombre d'étudiants, donc on ne peut pas choisir une salle sans l'inventer. Le cours n'est pas placé.",
}

ANOMALY_EXPLANATIONS: dict[str, str] = {
    "code salle préfixé <> normalisé": "CODE_SAL de la forme <>NNNN avec un nom « <Groupe>Salle (BAT-00) » : le préfixe est retiré car le nom correspond exactement à la salle NNNN d'EXP_SALLE.",
    "code salle préfixé <> rejeté": "CODE_SAL préfixé <> dont le code est absent d'EXP_SALLE ou dont le nom diffère de celui du catalogue : non rattaché à une salle, donc exclu de l'occupation par salle.",
    "salle absente du catalogue EXP_SALLE": "Le CODE_SAL du planning n'existe pas dans EXP_SALLE (souvent un code multiple de type <>). Les séances comptent dans les jours ouverts mais pas dans l'occupation d'une salle précise.",
    "promotion non reconnue dans EXP_PROMOTION": "Le NOM_DIP du planning n'est pas identique à un nom d'EXP_PROMOTION. Sans effet sur les calculs, qui utilisent EFFCALCU du planning.",
    "effectif inexploitable": "EFFCALCU vide, nul ou contradictoire pour cette séance. Elle est exclue des taux de remplissage et, si elle est dans un amphi LET, elle n'est pas reportée.",
    "effectif EFFCALCU supérieur à la capacité": "Effectif supérieur à la capacité de la salle dans EXP_SALLE : donnée à vérifier (capacité obsolète ou effectif cumulé). La séance est conservée telle quelle.",
    "ligne avec salle nommée mais sans CODE_SAL": "Une salle est nommée mais n'a pas de code : la séance ne peut pas être rattachée à une salle et n'est pas comptée.",
    "DUREE incohérente avec HDEBUT/HFIN": "La durée déclarée diffère de plus d'une minute de HFIN − HDEBUT. Le calcul utilise HDEBUT et HFIN.",
    "ligne multi-salles (@)": "La séance utilise plusieurs salles (séparées par @). Elle est ventilée par CODE_SAL.",
    "blocage administratif": "Réservation de 8 h ou plus, considérée comme un blocage et non un cours : exclue de l'occupation.",
}


def explain_anomaly(kind: str) -> str:
    for prefix, text in ANOMALY_EXPLANATIONS.items():
        if kind.startswith(prefix):
            return text
    return ""


def fr_number(value: float, digits: int = 1) -> str:
    return f"{value:,.{digits}f}".replace(",", " ").replace(".", ",")


def scenario_story(
    key: str,
    result: dict[str, Any],
    summary: dict[str, Any],
    source_events: list[dict[str, Any]],
    recipients: list[Room],
    reference_rows: list[dict[str, Any]],
    scenario_rows: list[dict[str, Any]],
    occupancy: dict[str, dict[str, dict[str, float]]],
    reference_conflicts: list[dict[str, Any]],
) -> list[str]:
    """Phrases simples et chiffrées qui expliquent ce que montre le tableau de bord pour un scénario."""
    flexible = key in {"H2", "H3b"}
    total = len(source_events)
    by_id = {event["id"]: event for event in source_events}
    smoothed = sum(
        1
        for source_id, status in result["statuses"].items()
        if status["status"] == "place"
        and (status["date"] != by_id[source_id]["date"] or status["start"] != by_id[source_id]["start"])
    )
    codes = {room.code for room in recipients}
    before = [occupancy["Référence"][code]["rate"] for code in codes if code in occupancy["Référence"]]
    after = [occupancy[key][code]["rate"] for code in codes if code in occupancy[key]]
    avg_before = sum(before) / len(before) if before else 0.0
    avg_after = sum(after) / len(after) if after else 0.0
    peak_code = max(codes & occupancy[key].keys(), key=lambda code: occupancy[key][code]["rate"])
    peak_name = next(room.name for room in recipients if room.code == peak_code)
    global_before = sum(row["rate"] for row in reference_rows) / max(1, len(reference_rows))
    global_after = sum(row["rate"] for row in scenario_rows) / max(1, len(scenario_rows))
    estimated = sum(
        1 for status in result["statuses"].values() if status.get("estimated") and status["status"] in {"place", "place_avec_chevauchement"}
    )
    not_moved = []
    if summary["unknownEffect"]:
        not_moved.append(f"{summary['unknownEffect']} car on ne connaît pas le nombre d'étudiants")
    if summary["noSolution"]:
        not_moved.append(f"{summary['noSolution']} faute d'horaire libre")
    if summary["uncovered"]:
        not_moved.append(f"{summary['uncovered']} car aucune salle n'est assez grande")
    lost = summary["unknownEffect"] + summary["noSolution"] + summary["uncovered"]
    if len(not_moved) == 1:
        not_moved = [not_moved[0].split(" ", 1)[1]]  # une seule raison : inutile de répéter le nombre
    unplaced_text = ""
    if lost:
        unplaced_text = f" {fr_number(lost, 0)} n'ont pas pu être déplacés" + (
            f", {not_moved[0]}." if len(not_moved) == 1 else " : " + ", ".join(not_moved) + "."
        )
    lines = [
        f"Il y a {fr_number(total, 0)} cours à déplacer ({fr_number(summary['sourceHours'], 0)} h). "
        f"{fr_number(summary['placed'], 0)} ont trouvé une salle." + unplaced_text
    ]
    if estimated:
        lines.append(
            f"Pour {estimated} cours, le nombre d'étudiants est une estimation (l'effectif de la promotion, qui est en général plus grand que le vrai). À vérifier."
        )
    if flexible:
        lines.append(
            f"{smoothed} cours ont été décalés (autre heure ou autre jour) parce que la salle était déjà prise. Les {summary['placed'] - smoothed} autres gardent leur horaire. "
            + ("Résultat : plus aucun cours en même temps dans la même salle." if not summary["conflicts"]
               else f"Il reste {summary['conflicts']} moments où deux cours ont lieu en même temps dans la même salle.")
        )
    else:
        lines.append(
            f"Aucun horaire ne change. Pour {summary['withConflicts']} cours, aucun amphi n'était libre : ils sont placés quand même, "
            f"ce qui crée {summary['conflicts']} moments où deux cours ont lieu en même temps dans la même salle (voir l'onglet Chevauchements)."
        )
    lines.append(
        f"Les salles qui reçoivent les cours sont occupées {fr_number(avg_before)} % du temps aujourd'hui, et {fr_number(avg_after)} % après le déplacement. "
        f"La plus chargée est {peak_name} ({fr_number(occupancy[key][peak_code]['rate'])} %). Les 3 amphis de Lettres tombent à 0 %."
    )
    lines.append(
        f"Le « taux d'occupation moyen » en haut du tableau ne change presque pas ({fr_number(global_before, 1)} % → {fr_number(global_after, 1)} %) : "
        "on ne fait que changer les cours de salle, il y a donc presque autant d'heures de cours au total. La petite baisse vient des cours qui n'ont pas pu être déplacés. "
        "Pour voir la différence, regardez les amphis de Droit (onglet Amphis) ou l'onglet Chevauchements."
    )
    return lines


def unknown_effect_rows(source_events: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Séances des amphis à reporter dont l'effectif reste inconnu (donc non placées dans les simulations)."""
    weekdays = ("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")
    rows = []
    for event in sorted(source_events, key=lambda e: (e["date"], e["start"], e["room_name"])):
        if event["effect"] is not None:
            continue
        group = " | ".join(event["diplomas"])
        origin = "Réservation sans promotion" if not group else ("Option UTLA" if group.startswith("<UTLA>") else "Autre groupe")
        rows.append(
            {
                "ID_SEANCE": event["id"], "DATE": event["date"].isoformat(), "JOUR": weekdays[event["date"].weekday()],
                "HDEBUT": f"{event['start'] // 60:02d}:{event['start'] % 60:02d}", "HFIN": f"{event['end'] // 60:02d}:{event['end'] % 60:02d}",
                "NOM_SAL": event["room_name"], "TYPE": event["type"], "LIBELLE_MAT": event["subject"], "NOM_DIP": group,
                "ORIGINE": origin, "HEURES": round(event["duration"] / 60, 2), "EFFECTIF_A_RENSEIGNER": "",
            }
        )
    return rows


def unknown_effect_html(created_at: str, rows: list[dict[str, Any]], estimated: int, total: int) -> str:
    """Page « Cours sans effectif » : pourquoi ces cours ne sont pas déplacés et la liste complète."""
    hours = sum(row["HEURES"] for row in rows)
    by_origin = Counter(row["ORIGINE"] for row in rows)
    by_room = Counter(row["NOM_SAL"] for row in rows)
    groups: dict[tuple[str, str], list[float]] = defaultdict(lambda: [0, 0.0])
    for row in rows:
        entry = groups[(row["NOM_DIP"] or "(aucune promotion)", row["LIBELLE_MAT"])]
        entry[0] += 1
        entry[1] += row["HEURES"]
    payload = json.dumps(
        {"rows": rows, "groups": [{"group": g, "subject": m, "n": v[0], "h": round(v[1], 1)} for (g, m), v in groups.items()]},
        ensure_ascii=False, separators=(",", ":"),
    ).replace("</", "<\\/")
    cards = "".join(
        f'<div class="kpi"><span>{safe_text(label)}</span><strong>{safe_text(value)}</strong></div>'
        for label, value in (
            ("Cours sans effectif", fr_number(len(rows), 0)),
            ("Heures concernées", f"{fr_number(hours, 0)} h"),
            ("Part des cours à déplacer", f"{fr_number(100 * len(rows) / max(1, total), 1)} %"),
            *[(origin, fr_number(count, 0)) for origin, count in by_origin.most_common()],
        )
    )
    rooms_text = ", ".join(f"{room} : {count}" for room, count in by_room.most_common())
    estimate_text = (
        f"{fr_number(estimated, 0)} autres cours ont reçu une estimation (l'effectif de leur promotion) et sont bien placés ; ils sont marqués « estimé » dans le fichier détail."
        if estimated else "L'option d'estimation n'est pas activée : aucun effectif n'a été estimé."
    )
    body = f"""<main><header><div><p class="eyebrow">UPPA · Direction du Patrimoine · SSH</p><h1>Cours sans effectif</h1><p class="muted">Créé le {safe_text(created_at)}</p></div></header>
<section class="card"><h2>De quoi s'agit-il ?</h2>
<p>Pour placer un cours dans une salle, il faut savoir <b>combien d'étudiants</b> y assistent. Dans le planning, ce nombre se calcule en comptant les identifiants des étudiants de chaque séance.</p>
<p>Pour les cours de cette page, le planning ne contient <b>ni liste d'étudiants, ni identifiant de promotion</b>. Le nombre d'étudiants est donc inconnu. Le programme <b>ne l'invente pas</b> : ces cours ne sont <b>pas déplacés</b> dans les simulations H1 à H3b, et leurs heures n'apparaissent donc pas dans les amphis de Droit.</p>
<p>{safe_text(estimate_text)}</p>
<p><b>Pour les placer :</b> renseigner le nombre d'étudiants dans la colonne « EFFECTIF_A_RENSEIGNER » du fichier CSV (bouton de téléchargement en haut de la page), ou refaire un export du planning qui contient la liste des étudiants de ces séances.</p></section>
<div class="kpis">{cards}</div>
<section class="card"><h2>Par amphi</h2><p>{safe_text(rooms_text or "Aucun.")}</p></section>
<section class="card"><h2>Par groupe</h2><div class="tablewrap"><table id="tg"><thead></thead><tbody></tbody></table></div></section>
<section class="card"><h2>Liste des cours</h2><div class="filters"><input id="q" type="search" placeholder="Rechercher (matière, groupe, amphi, date…)"><select id="o"><option value="">Toutes les origines</option></select></div><p class="muted" id="count"></p><div class="tablewrap"><table id="t"><thead></thead><tbody></tbody></table></div><p class="muted">Cliquez sur un titre de colonne pour trier.</p></section>
<script type="application/json" id="payload">{payload}</script><script>
(()=>{{const p=JSON.parse(document.getElementById('payload').textContent),$=s=>document.querySelector(s),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c])),fmt=n=>Number(n||0).toLocaleString('fr-FR',{{minimumFractionDigits:1,maximumFractionDigits:1}});
const cols=[['DATE','Date','t'],['JOUR','Jour','t'],['HDEBUT','Début','t'],['HFIN','Fin','t'],['NOM_SAL','Amphi','t'],['TYPE','Type','t'],['LIBELLE_MAT','Matière','t'],['NOM_DIP','Groupe','t'],['ORIGINE','Origine','t'],['HEURES','Heures','n']];
const st={{t:{{k:'DATE',d:1}},g:{{k:'n',d:-1}}}};
[...new Set(p.rows.map(r=>r.ORIGINE))].sort().forEach(v=>$('#o').add(new Option(v,v)));
function head(id,cs,s){{$(id+' thead').innerHTML='<tr>'+cs.map(([k,l,t])=>'<th data-k="'+k+'" class="'+(t==='n'?'num':'')+'">'+esc(l)+(k===st[s].k?(st[s].d>0?' ▲':' ▼'):' ↕')+'</th>').join('')+'</tr>';document.querySelectorAll(id+' th').forEach(th=>th.onclick=()=>{{const k=th.dataset.k;st[s].d=st[s].k===k?-st[s].d:(th.classList.contains('num')?-1:1);st[s].k=k;render()}})}}
const sorter=s=>(a,b)=>{{const x=a[st[s].k],y=b[st[s].k];return(typeof x==='string'?x.localeCompare(y,'fr'):x-y)*st[s].d}};
function render(){{const q=$('#q').value.trim().toLowerCase(),o=$('#o').value;const v=p.rows.filter(r=>(!o||r.ORIGINE===o)&&(!q||Object.values(r).join(' ').toLowerCase().includes(q))).sort(sorter('t'));
head('#t',cols,'t');$('#t tbody').innerHTML=v.map(r=>'<tr>'+cols.map(([k,l,t])=>'<td'+(t==='n'?' class="num"':'')+'>'+esc(t==='n'?fmt(r[k]):r[k])+'</td>').join('')+'</tr>').join('');$('#count').textContent=v.length+' cours affiché(s) sur '+p.rows.length+'.';
const gc=[['group','Groupe','t'],['subject','Matière','t'],['n','Nombre de cours','n'],['h','Heures','n']];head('#tg',gc,'g');$('#tg tbody').innerHTML=p.groups.slice().sort(sorter('g')).map(r=>'<tr>'+gc.map(([k,l,t])=>'<td'+(t==='n'?' class="num"':'')+'>'+esc(t==='n'?(k==='h'?fmt(r[k]):r[k]):r[k])+'</td>').join('')+'</tr>').join('')}}
$('#q').addEventListener('input',render);$('#o').addEventListener('change',render);render();}})();
</script></main>"""
    return html_document("Cours sans effectif", created_at, body, BASE_CSS)


def amphis_html(
    created_at: str,
    rooms: dict[str, Room],
    building_sites: dict[str, dict[str, str]],
    streams: dict[str, list[dict[str, Any]]],
    active_dates: list[dt.date],
    exam_periods: list[tuple[dt.date, dt.date]] | None = None,
) -> str:
    """Page dédiée aux amphithéâtres : occupation avant/après dans chaque scénario, filtrable et triable."""
    amphis = sorted((r for r in rooms.values() if r.kind == "amphi"), key=lambda r: (r.building, r.name.casefold()))
    exam_periods = exam_periods or []
    exam_dates = [date for date in active_dates if is_exam_date(date, exam_periods)]
    exam_set = set(exam_dates)
    date_sets = {"AN": active_dates, "EX": exam_dates, "HX": [date for date in active_dates if date not in exam_set]}
    occupancy = {
        period: {name: occupancy_by_room(stream, amphis, dates) for name, stream in streams.items()}
        for period, dates in date_sets.items()
    }
    names = list(streams)
    exam_text = " ; ".join(f"{a.strftime('%d/%m/%Y')} – {b.strftime('%d/%m/%Y')}" for a, b in exam_periods) or "aucune"
    data = []
    for room in amphis:
        site = building_sites.get(room.building, {})
        row: dict[str, Any] = {
            "name": room.name, "code": room.code, "building": room.building, "city": site.get("ville", ""),
            "site": site.get("site", ""), "capacity": room.capacity or 0,
            "removed": room.building == "LET", "host": room.building == "DEG",
        }
        row["P"] = {
            period: {
                **{f"h_{name}": round(occupancy[period][name][room.code]["hours"], 1) for name in names},
                **{f"r_{name}": round(occupancy[period][name][room.code]["rate"], 2) for name in names},
            }
            for period in date_sets
        }
        data.append(row)
    payload = json.dumps({"rows": data, "scenarios": names}, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    body = f"""<main><header><div><p class="eyebrow">UPPA · Direction du Patrimoine · SSH</p><h1>Amphithéâtres</h1><p class="muted">Occupation de chaque amphi avant (Référence) et après report. Créé le {safe_text(created_at)}</p></div></header>
<div class="filters"><div class="msel" id="m-city"></div><div class="msel" id="m-site"></div><div class="msel" id="m-bld"></div><input id="f-q" type="search" placeholder="Rechercher un amphi…"><div class="seg"><button data-scope="all" aria-pressed="true">Tous les amphis</button><button data-scope="let" aria-pressed="false">LET (supprimés)</button><button data-scope="deg" aria-pressed="false">DEG (accueil)</button></div><div class="pseg" role="group" aria-label="Période"><button data-period="AN" aria-pressed="true">Année</button><button data-period="EX" aria-pressed="false" title="{safe_text(exam_text)}">Examens</button><button data-period="HX" aria-pressed="false">Hors examens</button></div><button type="button" class="reset" id="f-reset">Réinitialiser</button></div>
<p class="note" id="count"></p>
<p class="muted" id="period-note"></p>
<section class="card" id="avg" style="margin-bottom:12px"></section>
<div class="tablewrap"><table id="t"><thead></thead><tbody></tbody></table></div>
<p class="muted">Cliquez sur un titre de colonne pour trier (nouveau clic : ordre inverse). Taux = heures occupées entre 08:00 et 18:00 divisées par les heures théoriques des {len(active_dates)} jours pédagogiques. Les lignes vertes sont les amphis LET supprimés dans les scénarios (taux 0 après report) ; les lignes ambrées sont les amphis DEG d'accueil. Un taux supérieur à 100 % signale des chevauchements conservés.</p>
<script type="application/json" id="payload">{payload}</script><script>
(()=>{{const p=JSON.parse(document.getElementById('payload').textContent),rows=p.rows,sc=p.scenarios,$=s=>document.querySelector(s),fmt=(n,d=1)=>Number(n||0).toLocaleString('fr-FR',{{minimumFractionDigits:d,maximumFractionDigits:d}}),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c]));
let scope='all',key='r_'+sc[0],dir=-1;const chk=new Set();
const cols=[['name','Amphi','t'],['building','Bât.','t'],['city','Ville','t'],['site','Site','t'],['capacity','Places','n'],...sc.flatMap(n=>[['r_'+n,'Taux '+n+' (%)','n'],['h_'+n,'Heures '+n,'n']]),...sc.slice(1).map(n=>['d_'+n,'Écart '+n+' vs Réf. (pts)','n'])];
const days={{AN:{len(date_sets["AN"])},EX:{len(date_sets["EX"])},HX:{len(date_sets["HX"])}}},pname={{AN:"l'année complète",EX:"les périodes d'examen ({safe_text(exam_text)})",HX:"les jours hors périodes d'examen"}};let per='AN';
function setPeriod(q){{per=q;rows.forEach(r=>{{Object.assign(r,r.P[q]);sc.slice(1).forEach(n=>r['d_'+n]=Math.round((r['r_'+n]-r['r_'+sc[0]])*100)/100)}});$('#period-note').textContent='Période affichée : '+pname[q]+' — '+days[q]+' jour(s) pédagogique(s).'+(days[q]===0?' Aucune donnée sur cette période dans l’export.':'');document.querySelectorAll('.pseg button').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.period===q)))}}
setPeriod('AN');
const sel={{city:new Set(),site:new Set(),building:new Set()}},labels={{city:['Toutes les villes','ville(s)'],site:['Tous les sites','site(s)'],building:['Tous les bâtiments','bâtiment(s)']}},ids={{city:'#m-city',site:'#m-site',building:'#m-bld'}};
function base(){{return rows.filter(r=>scope==='all'||(scope==='let'&&r.removed)||(scope==='deg'&&r.host))}}
function pool(field){{let p=base();if(field!=='city'&&sel.city.size)p=p.filter(r=>sel.city.has(r.city));if(field==='building'&&sel.site.size)p=p.filter(r=>sel.site.has(r.site));return p}}
function msel(field){{const root=$(ids[field]),open=root.dataset.open==='1',counts={{}};pool(field).forEach(r=>{{if(r[field])counts[r[field]]=(counts[r[field]]||0)+1}});[...sel[field]].forEach(v=>{{if(!(v in counts))sel[field].delete(v)}});
const n=sel[field].size,label=n?n+' '+labels[field][1]+' sélectionné(s)':labels[field][0];
root.innerHTML='<button type="button" class="msel-btn" aria-expanded="'+open+'">'+esc(label)+' ▾</button><div class="msel-panel"'+(open?'':' hidden')+'><div class="msel-act"><a href="#" data-a="all">Tout cocher</a> · <a href="#" data-a="none">Tout décocher</a></div>'+Object.keys(counts).sort().map(v=>'<label><input type="checkbox" value="'+esc(v)+'"'+(sel[field].has(v)?' checked':'')+'> '+esc(v)+' <small>('+counts[v]+')</small></label>').join('')+'</div>';
root.querySelector('.msel-btn').onclick=e=>{{e.stopPropagation();const o=root.dataset.open==='1';document.querySelectorAll('.msel').forEach(m=>{{m.dataset.open='0';m.querySelector('.msel-panel').hidden=true}});root.dataset.open=o?'0':'1';root.querySelector('.msel-panel').hidden=o}};
root.querySelector('.msel-panel').onclick=e=>e.stopPropagation();
root.querySelectorAll('input').forEach(c=>c.onchange=()=>{{c.checked?sel[field].add(c.value):sel[field].delete(c.value);render()}});
root.querySelectorAll('[data-a]').forEach(a=>a.onclick=e=>{{e.preventDefault();sel[field].clear();if(a.dataset.a==='all')Object.keys(counts).forEach(v=>sel[field].add(v));render()}});}}
function visible(){{const q=$('#f-q').value.trim().toLowerCase();return base().filter(r=>(!sel.city.size||sel.city.has(r.city))&&(!sel.site.size||sel.site.has(r.site))&&(!sel.building.size||sel.building.has(r.building))&&(!q||(r.name+' '+r.code+' '+r.building).toLowerCase().includes(q)))}}
function render(){{['city','site','building'].forEach(msel);
const v=visible().sort((a,b)=>{{const x=a[key],y=b[key];return(typeof x==='string'?x.localeCompare(y,'fr'):(x-y))*dir||a.name.localeCompare(b.name,'fr')}});
$('#t thead').innerHTML='<tr><th><input type="checkbox" id="chk-all" title="Cocher / décocher les amphis affichés" '+(v.length&&v.every(r=>chk.has(r.code))?'checked':'')+'></th>'+cols.map(([k,l,t])=>'<th data-k="'+k+'" class="'+(t==='n'?'num':'')+'">'+esc(l)+(k===key?(dir>0?' ▲':' ▼'):' ↕')+'</th>').join('')+'</tr>';
$('#t tbody').innerHTML=v.map(r=>'<tr'+(r.removed?' style="background:#e6f3ef"':r.host?' style="background:#fff5e0"':'')+'><td><input type="checkbox" class="rc" value="'+esc(r.code)+'"'+(chk.has(r.code)?' checked':'')+'></td>'+cols.map(([k,l,t])=>t==='n'?'<td class="num">'+(k==='capacity'?r[k]:fmt(r[k],k[0]==='h'?1:2))+'</td>':'<td>'+esc(r[k])+'</td>').join('')+'</tr>').join('');
$('#count').textContent=v.length+' amphi(s) affiché(s) sur '+rows.length+'.';
document.querySelectorAll('#t th[data-k]').forEach(th=>th.onclick=()=>{{dir=key===th.dataset.k?-dir:(th.classList.contains('num')?-1:1);key=th.dataset.k;render()}});
document.querySelectorAll('#t .rc').forEach(c=>c.onchange=()=>{{c.checked?chk.add(c.value):chk.delete(c.value);render()}});
const all=$('#chk-all');if(all)all.onchange=()=>{{v.forEach(r=>all.checked?chk.add(r.code):chk.delete(r.code));render()}};
average(v);}}
function average(v){{const picked=rows.filter(r=>chk.has(r.code)),set=picked.length?picked:v,n=set.length,own=picked.length>0;
const mean=f=>n?set.reduce((a,r)=>a+r[f],0)/n:0,sum=f=>set.reduce((a,r)=>a+r[f],0),ref=mean('r_'+sc[0]);
$('#avg').innerHTML='<h2>Moyenne générale des amphis '+(own?'sélectionnés ('+n+' coché'+(n>1?'s':'')+')':'affichés ('+n+') — cochez des amphis pour choisir')+'</h2>'+(n?'<div class="tablewrap"><table><thead><tr><th>Cas</th><th class="num">Taux moyen (%)</th><th class="num">Écart vs Référence (pts)</th><th class="num">Heures totales</th><th class="num">Heures moyennes / amphi</th></tr></thead><tbody>'+sc.map(s=>{{const m=mean('r_'+s),d=m-ref;return '<tr><td>'+esc(s)+'</td><td class="num"><b>'+fmt(m,2)+'</b></td><td class="num">'+(s===sc[0]?'—':(d>0?'+':'')+fmt(d,2))+'</td><td class="num">'+fmt(sum('h_'+s),1)+'</td><td class="num">'+fmt(sum('h_'+s)/n,1)+'</td></tr>'}}).join('')+'</tbody></table></div><p class="muted">Moyenne arithmétique des taux des amphis retenus (tous ont le même nombre de jours théoriques, donc identique à la moyenne pondérée par les heures). '+(own?'<a href="#" id="clr">Décocher tous</a>':'')+'</p>':'<p class="muted">Aucun amphi affiché.</p>');
const c=$('#clr');if(c)c.onclick=e=>{{e.preventDefault();chk.clear();render()}};}}
document.addEventListener('click',()=>document.querySelectorAll('.msel').forEach(m=>{{if(m.dataset.open==='1'){{m.dataset.open='0';m.querySelector('.msel-panel').hidden=true}}}}));
$('#f-q').addEventListener('input',render);
$('#f-reset').addEventListener('click',()=>{{Object.values(sel).forEach(x=>x.clear());chk.clear();$('#f-q').value='';scope='all';setPeriod('AN');document.querySelectorAll('.seg button').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.scope==='all')));render()}});
document.querySelectorAll('.pseg button').forEach(b=>b.addEventListener('click',()=>{{setPeriod(b.dataset.period);render()}}));
document.querySelectorAll('.seg button').forEach(b=>b.addEventListener('click',()=>{{scope=b.dataset.scope;document.querySelectorAll('.seg button').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));render()}}));
render();}})();
</script></main>"""
    extra_css = ".pseg{display:flex}.pseg button{font:inherit;background:#fff;border:1px solid var(--line);padding:7px 10px;cursor:pointer}.pseg button+button{border-left:0}.pseg button[aria-pressed=true]{background:#dcebe7;color:#14594e;font-weight:700}.msel{position:relative}.msel-btn,.reset{font:inherit;background:#fff;border:1px solid var(--line);border-radius:5px;padding:7px 10px;cursor:pointer;min-width:150px;text-align:left}.reset{min-width:0;color:var(--muted)}.msel-panel{position:absolute;z-index:5;top:100%;left:0;margin-top:3px;min-width:240px;max-height:300px;overflow:auto;background:#fff;border:1px solid var(--line);border-radius:6px;box-shadow:0 6px 18px rgba(0,0,0,.12);padding:6px 8px}.msel-panel[hidden]{display:none}.msel-act{font-size:12px;padding:3px 2px 6px;border-bottom:1px solid var(--line);margin-bottom:4px}.msel-panel label{display:flex;align-items:center;gap:7px;padding:4px 2px;font-size:13px;text-transform:none;color:var(--ink);cursor:pointer}.msel-panel label:hover{background:#f1f6f5}.msel-panel input{padding:0}.msel-panel small{color:var(--muted)}"
    return html_document("Amphithéâtres UPPA", created_at, body, BASE_CSS + extra_css)


def combined_template_dashboard(
    created_at: str,
    stamp: str,
    pages: dict[str, str],
    oversize_csv_name: str,
    detail_csv_name: str,
    stories: dict[str, list[str]] | None = None,
    unknown_csv_name: str | None = None,
) -> str:
    payload = json.dumps(pages, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    case_payload = json.dumps(
        {"views": CASE_EXPLANATIONS, "status": STATUS_EXPLANATIONS, "live": stories or {}}, ensure_ascii=False, separators=(",", ":")
    ).replace("</", "<\\/")
    buttons = "".join(
        f'<button type="button" class="view-tab{" active" if index == 0 else ""}" data-view="{safe_text(key)}">{safe_text(label)}</button>'
        for index, (key, label) in enumerate(
            (("amphis", "Amphis"), ("sans_effectif", "Cours sans effectif"), ("reference", "Référence"), ("H1", "H1"), ("H2", "H2"), ("H3a", "H3a"), ("H3b", "H3b"), ("surdimensionnement", "Surdimensionnement"))
        )
    )
    body = f"""<main><header><div><p>UPPA · Direction du Patrimoine · SSH</p><h1>Occupation des salles · dossier complet</h1><span>Référence 2025–2026 et simulations H1, H2, H3a, H3b</span></div><div class="stamp">Créé le {safe_text(created_at)}</div></header>
<nav aria-label="Vues du dossier">{buttons}</nav><div class="downloads"><a href="{safe_text(oversize_csv_name)}" download>↓ CSV surdimensionnement</a><a href="{safe_text(detail_csv_name)}" download>↓ CSV détail des séances</a>{f'<a href="{safe_text(unknown_csv_name)}" download>↓ CSV cours sans effectif</a>' if unknown_csv_name else ''}</div>
<section id="case-info" class="case-info" aria-live="polite"></section>
<iframe id="view" title="Tableau de bord d'occupation" sandbox="allow-scripts" srcdoc=""></iframe>
<script type="application/json" id="case-texts">{case_payload}</script><script type="application/json" id="embedded-pages">{payload}</script><script>
(()=>{{const pages=JSON.parse(document.getElementById('embedded-pages').textContent),cases=JSON.parse(document.getElementById('case-texts').textContent),frame=document.getElementById('view'),info=document.getElementById('case-info');function esc(t){{return String(t).replace(/[&<>"']/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c]))}}let panelOpen=true;function explain(key){{const c=cases.views[key];if(!c){{info.hidden=true;return}}info.hidden=false;let h='<details id="ci-main"'+(panelOpen?' open':'')+'><summary><h2>'+esc(c.title)+'</h2></summary><p><strong>En clair :</strong> '+esc(c.what)+'</p><p><strong>Pour bien lire :</strong> '+esc(c.how)+'</p>';if(cases.live[key]){{h+='<p><strong>Les chiffres de ce cas :</strong></p><ul>'+cases.live[key].map(t=>'<li>'+esc(t)+'</li>').join('')+'</ul>'}}if(/^H/.test(key)){{h+='<details><summary>Que veulent dire les mentions « placé », « sans solution »… (fichier détail) ?</summary><ul>'+Object.entries(cases.status).map(([k,v])=>'<li><code>'+esc(k)+'</code> : '+esc(v)+'</li>').join('')+'</ul></details>'}}h+='</details>';info.innerHTML=h;const d=document.getElementById('ci-main');d.addEventListener('toggle',()=>{{panelOpen=d.open}})}}function show(key){{explain(key);frame.srcdoc=pages[key];document.querySelectorAll('.view-tab').forEach(button=>button.classList.toggle('active',button.dataset.view===key));}}document.querySelectorAll('.view-tab').forEach(button=>button.addEventListener('click',()=>show(button.dataset.view)));show(document.querySelector('.view-tab').dataset.view);}})();
</script></main>"""
    css = """
:root{color-scheme:light;--ink:#17252a;--muted:#52636b;--line:#d1dadb;--paper:#eef2f1;--teal:#17675d}*{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);font:14px/1.45 system-ui,'Segoe UI',sans-serif}main{padding:18px 20px 20px}header{display:flex;justify-content:space-between;align-items:end;border-bottom:1px solid var(--line);padding-bottom:12px}header p{font:600 11px monospace;color:var(--teal);text-transform:uppercase;margin:0}h1{font:600 26px Georgia,serif;margin:5px 0}header span,.stamp{color:var(--muted);font-size:12px}.stamp{text-align:right;font-family:monospace}nav{display:flex;gap:4px;overflow:auto;border-bottom:1px solid var(--line);margin-top:10px}.view-tab{border:0;border-bottom:2px solid transparent;background:none;padding:9px 14px;color:var(--muted);font:600 13px system-ui;cursor:pointer;white-space:nowrap}.view-tab.active{color:var(--teal);border-color:var(--teal)}.downloads{display:flex;justify-content:flex-end;gap:18px;padding:8px 2px;font-size:12px}.downloads a{color:var(--teal)}.case-info{background:#fff;border:1px solid var(--line);border-left:4px solid var(--teal);border-radius:6px;padding:10px 16px;margin:4px 0 10px}.case-info h2{font:600 16px Georgia,serif;margin:0;display:inline}.case-info>details>summary{cursor:pointer;list-style:none;display:flex;align-items:center;gap:8px;padding:2px 0}.case-info>details>summary::-webkit-details-marker{display:none}.case-info>details>summary::before{content:'\\25B6';font-size:11px;color:var(--teal);transition:transform .15s}.case-info>details[open]>summary::before{transform:rotate(90deg)}.case-info>details>summary:hover h2{text-decoration:underline}.case-info>details[open]>summary{margin-bottom:6px}.case-info p{margin:4px 0;font-size:13px}.case-info summary{cursor:pointer;color:var(--teal);font-size:12px}.case-info li{font-size:12px;margin:3px 0}.case-info code{background:#e7eeee;padding:1px 4px}iframe{display:block;width:100%;height:calc(100vh - 165px);min-height:550px;border:0;background:var(--paper)}@media(max-width:700px){main{padding:12px}header{align-items:start;flex-direction:column}.stamp{text-align:left}.downloads{justify-content:flex-start;flex-wrap:wrap}iframe{height:calc(100vh - 205px);min-height:480px}}
"""
    return html_document("Occupation des salles UPPA · dossier complet", created_at, body, css)


def summary_html(
    created_at: str,
    stamp: str,
    data: dict[str, Any],
    scenarios: dict[str, dict[str, Any]],
    reference_total_hours: float,
    saturated: list[dict[str, Any]],
    calendars: dict[str, Any],
    manual_checks: list[dict[str, Any]],
) -> str:
    table_rows = "".join(
        "<tr><td>{}</td><td class='num'>{}</td><td class='num'>{}</td><td class='num'>{}</td><td class='num'>{}</td><td class='num'>{}</td><td class='num'>{}</td></tr>".format(
            safe_text(label),
            summary["placed"],
            summary["withConflicts"],
            summary["uncovered"],
            summary["noSolution"],
            summary["unknownEffect"],
            f"{summary['sourceHours']:.2f} / {summary['accountedHours']:.2f}",
        )
        for label, summary in scenarios.items()
    )
    saturated_rows = "".join(
        f"<tr><td>{safe_text(item['room'])}</td><td>{safe_text(item['building'])}</td><td>{item['year']} · S{item['week']}</td><td class='num'>{item['rate']:.2f}%</td></tr>"
        for item in saturated[:60]
    ) or "<tr><td colspan='4'>Aucune salle/semaine au-dessus de 100 %.</td></tr>"
    promotions = "".join(f"<li>{safe_text(name)}</li>" for name in data["unrecognized_promotions"][:150])
    extra_promotions = max(0, len(data["unrecognized_promotions"]) - 150)
    source_lines = "".join(
        f"<tr><td>{safe_text(name)}</td><td class='num'>{value}</td></tr>"
        for name, value in sorted(data["blank_counts"].items(), key=lambda item: item[0].casefold())
    )
    manual_rows = "".join(
        f"<tr><td>{safe_text(row['code'])}</td><td class='num'>{row['rawEvents']}</td>"
        f"<td class='num'>{row['rawHours']:.2f}</td><td class='num'>{row['engineEvents']}</td>"
        f"<td class='num'>{row['engineHours']:.2f}</td><td class='num'>{row['differenceMinutes']}</td></tr>"
        for row in manual_checks
    )
    notes = [
        f"Le planning comporte {data['source_rows']} lignes et {len(data['events'])} séances physiques dédoublonnées par CODE_SAL/date/heure/matière/type.",
        f"PERIODE : {len(data['period_values'])} valeurs distinctes, exemple {', '.join(data['period_values'][:5])} ; {data['period_range_rows']} lignes ont une plage de plusieurs semaines. Chaque DDEBUT a été contrôlé contre PERIODE (écarts : {data['iso_mismatch']}) ; jours JOUR incohérents : {data['day_mismatch']}.",
        f"DUREE déclarée : {data['duration_min']} à {data['duration_max']} h ; écarts supérieurs à une minute avec HDEBUT/HFIN : {data['duration_mismatch']}. Blocages d'au moins 8 h retirés.",
        f"Occurrences multi-salles marquées @ : NOM_SAL {data['at_name_count']} lignes ; CODE_SAL {data['at_code_count']} lignes.",
        "EFFCALCU est pris comme effectif de la séance. Les lignes de promotions liées qui répètent le même EFFCALCU ne sont jamais additionnées.",
        "Les dates d'ouverture pédagogique sont les jours ouvrés où l'export contient au moins une séance non bloquante. Les jours fériés nationaux français sont retranchés ; les périodes de fermeture non présentes dans les données sont corrigeables avec --exclure-semaines.",
        "Les salles d'accueil DEG et les grandes salles d'accueil gardent leurs séances de référence ; seuls les enseignements des amphithéâtres LET sont déplacés.",
        "Les taux ne comptent que les minutes dans la plage 08:00–18:00. La conservation des heures de transfert porte sur la durée complète HDEBUT/HFIN. Le lissage est une recherche théorique à la minute, dans les jours pédagogiques observés, entre 08:00 et 18:00 ; il ne constitue pas un planning validé.",
        f"Semaines exclues explicitement par option : {data['excluded_week_count']} lignes de séances ; calendrier fondé sur les jours ouvrés observés dans l'export.",
        "Les capacités d'examen en configuration plateau ne sont pas modélisées ; les usages d'examen présents dans l'export ne décrivent pas une implantation de plateau.",
        f"Blocages administratifs de durée >= 8 h exclus : {data['admin_count']}. Séances sur jour férié exclues : {data['holiday_count']}. Séances pendant les fermetures de l'université (Toussaint, Noël, hiver, printemps) exclues : {data['closure_count']}. Samedis/dimanches hors base ouvrée : {data['weekend_count']}.",
        f"Séances sans effectif exploitable dans tout le planning localisé : {data['missing_effect']} ; événements source des amphithéâtres LET : {data['missing_effect_source']}.",
        f"Salle renseignée sans CODE_SAL : {data['missing_room_code_count']} lignes ; codes absents du catalogue : {sum(data['unknown_codes'].values())} lignes.",
        f"Durées HDEBUT/HFIN qui diffèrent de DUREE de plus d'une minute : {data['duration_mismatch']}.",
    ]
    list_notes = "".join(f"<li>{safe_text(note)}</li>" for note in notes)
    body = f"""<main><header><p class="eyebrow">UPPA · Direction du Patrimoine · SSH</p><h1>Synthèse · Faisabilité du report des amphithéâtres LET</h1><p>Créé le {safe_text(created_at)} · Période analysée {calendars['start']} au {calendars['end']}</p></header>
<section><h2>Indicateurs vérifiables</h2><p>Heures-salles de référence : <strong>{reference_total_hours:.2f} h</strong>. {len(data['events'])} séances localisées après dédoublonnage ; {data['source_rows']} lignes brutes.</p><p>Le modèle de simulation n'affirme une faisabilité que pour les séances effectivement placées dans une salle de capacité suffisante. « Effectif inconnu », « non couvert » et « sans solution » restent des résultats distincts.</p></section>
<section><h2>Scénarios</h2><div class="tablewrap"><table><thead><tr><th>Scénario</th><th>Placées</th><th>Avec chevauchement</th><th>Non couvertes</th><th>Sans solution</th><th>Effectif inconnu</th><th>Heures source / comptabilisées</th></tr></thead><tbody>{table_rows}</tbody></table></div></section>
<section><h2>Saturation hebdomadaire de référence</h2><p>Somme des heures réservées par local divisée par les heures ouvrées de la semaine. Les taux supérieurs à 100 % signalent des chevauchements conservés, pas une capacité physique supérieure.</p><div class="tablewrap"><table><thead><tr><th>Local</th><th>Bât.</th><th>Semaine</th><th>Taux</th></tr></thead><tbody>{saturated_rows}</tbody></table></div></section>
<section><h2>Hypothèses et limites</h2><ul>{list_notes}</ul></section>
<section><h2>Contrôle qualité des sources</h2><div class="tablewrap"><table><thead><tr><th>Colonne</th><th>Valeurs vides</th></tr></thead><tbody>{source_lines}</tbody></table></div><p>PERIODE et la date sont contrôlées à l'exécution. Les occupations impossibles à attribuer à une salle physique (CODE_SAL manquant/inconnu) ne sont pas réparties arbitrairement.</p><details><summary>Promotions NOM_DIP non reconnues à l'identique par EXP_PROMOTION</summary><ul>{promotions}</ul>{f'<p>Liste tronquée à 150 valeurs ; {extra_promotions} autres.</p>' if extra_promotions else ''}</details><p>Effectif noté comme inconnu si EFFCALCU est vide, nul, non entier ou contradictoire sur une même clé de séance.</p></section>
<section><h2>Rapprochement indépendant</h2><p>Recomptage direct des intervalles dans le planning brut pour les amphithéâtres LET, comparé aux séances dédoublonnées du moteur. Les heures présentées sont limitées à 08:00–18:00 ; l'écart attendu est nul.</p><div class="tablewrap"><table><thead><tr><th>CODE_SAL</th><th>Occurrences brutes uniques</th><th>Heures brutes 08–18</th><th>Occurrences moteur</th><th>Heures moteur 08–18</th><th>Écart (minutes)</th></tr></thead><tbody>{manual_rows}</tbody></table></div></section>
<footer>Les heures et capacités sont issues des CSV livrés ; aucune valeur d'effectif n'est extrapolée. HTML autonome · génération {safe_text(stamp)}.</footer></main>"""
    return html_document("Synthèse occupation UPPA", created_at, body, BASE_CSS)


def overdimension_rows(events: list[dict[str, Any]], rooms: dict[str, Room]) -> list[dict[str, Any]]:
    allowed_types = {"cours", "cm", "td", "ctd"}
    eligible = [
        event
        for event in events
        if event["type"].strip().casefold() in allowed_types
        and event["room_kind"] in {"amphi", "salle_cours", "grande_salle"}
        and event["effect"] is not None
        and event["capacity"] is not None
        and event["capacity"] >= 2 * event["effect"]
        and event["capacity"] - event["effect"] >= 20
    ]
    series: dict[tuple[str, str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for event in eligible:
        series[(event["room_code"], event["subject"], event["type"], "|".join(event["diplomas"]))].append(event)
    source_groups = room_events(events)
    output: list[dict[str, Any]] = []
    for series_key, occurrences in series.items():
        source_code = series_key[0]
        max_effect = max(event["effect"] for event in occurrences)
        source_building = occurrences[0]["building"]
        choices = [
            room
            for room in rooms.values()
            if room.code != source_code and room.capacity is not None and room.capacity >= max_effect
        ]
        choices.sort(key=lambda room: (room.building != source_building, room.capacity, room.code))
        replacement = None
        for room in choices:
            free_all = True
            for occurrence in occurrences:
                probe = {"date": occurrence["date"], "start": occurrence["start"], "end": occurrence["end"]}
                if any(overlap(probe, existing) for existing in source_groups.get((room.code, occurrence["date"]), [])):
                    free_all = False
                    break
            if free_all:
                replacement = room
                break
        for occurrence in occurrences:
            output.append(
                {
                    "date": occurrence["date"].isoformat(),
                    "room": occurrence["room_name"],
                    "code": occurrence["room_code"],
                    "building": occurrence["building"],
                    "capacity": occurrence["capacity"],
                    "effect": occurrence["effect"],
                    "unused": occurrence["capacity"] - occurrence["effect"],
                    "subject": occurrence["subject"],
                    "type": occurrence["type"],
                    "start": occurrence["start"],
                    "end": occurrence["end"],
                    "diplomas": ", ".join(occurrence["diplomas"]),
                    "seriesOccurrences": len(occurrences),
                    "replacement": replacement.name if replacement else "Aucun local assez grand libre à toutes les occurrences",
                    "replacementCode": replacement.code if replacement else "",
                }
            )
    return sorted(output, key=lambda row: (row["date"], row["room"], row["start"], row["subject"]))


def overdimension_html(rows: list[dict[str, Any]], created_at: str) -> str:
    headers = ("Date", "Local", "Bât.", "Capacité", "Effectif EFFCALCU", "Places inutilisées", "Type", "Matière", "Promotions", "Occurrences série", "Plus petite salle libre à toutes les occurrences")
    body_rows = []
    for row in rows:
        body_rows.append(
            "<tr>" + "".join(
                f"<td>{safe_text(value)}</td>"
                for value in (
                    row["date"], row["room"], row["building"], row["capacity"], row["effect"], row["unused"],
                    row["type"], row["subject"], row["diplomas"], row["seriesOccurrences"], row["replacement"],
                )
            ) + "</tr>"
        )
    rows_html = "".join(body_rows) or f"<tr><td colspan='{len(headers)}'>Aucun cas conforme au seuil.</td></tr>"
    table_head = "".join(f"<th>{safe_text(label)}</th>" for label in headers)
    body = f"""<main><header><p class="eyebrow">UPPA · Contrôle des capacités</p><h1>Surdimensionnement des locaux</h1><p>Créé le {safe_text(created_at)} · {len(rows)} occurrences détectées</p></header><p>TYPE ∈ Cours, CM, TD, CTD ; capacité ≥ 2 × EFFCALCU et au moins 20 places inutilisées. Le local suggéré est le plus petit local suffisant, priorisé dans le même bâtiment, libre sur toutes les occurrences de la série.</p><div class="tablewrap"><table id="oversized"><thead><tr>{table_head}</tr></thead><tbody>{rows_html}</tbody></table></div></main><script>document.querySelectorAll('th').forEach((th,i)=>th.onclick=()=>{{let t=th.closest('table'),b=t.tBodies[0],r=[...b.rows],d=th.dataset.d==='a'?-1:1;th.dataset.d=d===1?'a':'d';r.sort((x,y)=>x.cells[i].innerText.localeCompare(y.cells[i].innerText,'fr',{{numeric:true}})*d);r.forEach(x=>b.appendChild(x));}})</script>"""
    return html_document("Surdimensionnement UPPA", created_at, body, BASE_CSS)


def write_csv(path: Path, headers: list[str], rows: list[dict[str, Any]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=headers, delimiter=";", lineterminator="\n", extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def detail_rows(
    events: list[dict[str, Any]], scenario_results: dict[str, dict[str, Any]], rooms: dict[str, Room] | None = None
) -> list[dict[str, Any]]:
    rooms = rooms or {}
    details: list[dict[str, Any]] = []
    for event in events:
        item: dict[str, Any] = {
            "ID_SEANCE": event["id"],
            "DATE": event["date"].isoformat(),
            "PERIODE": event["period"],
            "JOUR": ("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")[event["date"].weekday()],
            "HDEBUT": f"{event['start'] // 60:02d}:{event['start'] % 60:02d}",
            "HFIN": f"{event['end'] // 60:02d}:{event['end'] % 60:02d}",
            "DUREE_HEURES": f"{event['duration'] / 60:.4f}",
            "TYPE": event["type"],
            "LIBELLE_MAT": event["subject"],
            "NOM_DIP": " | ".join(event["diplomas"]),
            "NOM_SAL": event["room_name"],
            "CODE_SAL": event["room_code"],
            "CAPA": event["capacity"],
            "EFFCALCU": event["effect"] if event["effect"] is not None else "",
            "ECART_PLACES": event["capacity"] - event["effect"] if event["effect"] is not None and event["capacity"] is not None else "",
            "STATUT_EFFECTIF": event["effect_issue"],
            "LIGNES_SOURCE": event["source_rows"],
        }
        for scenario, result in scenario_results.items():
            status = result["statuses"].get(event["id"])
            item[f"{scenario}_STATUT"] = status["status"] if status else "non concerné"
            item[f"{scenario}_LOCAL"] = status["destination"] if status else ""
            item[f"{scenario}_CODE"] = status["destination_code"] if status else ""
            item[f"{scenario}_DATE"] = status["date"].isoformat() if status and status["status"] == "place" else ""
            item[f"{scenario}_DEBUT"] = f"{status['start'] // 60:02d}:{status['start'] % 60:02d}" if status and status["status"] == "place" else ""
            item[f"{scenario}_MOTIF"] = status["reason"] if status else ""
            item[f"{scenario}_EFFECTIF_ESTIME"] = "oui" if status and status.get("estimated") else ""
            item[f"{scenario}_STATUT_FINAL"] = final_status(event, status)
            destination = rooms.get(status["destination_code"]) if status and status["destination_code"] else None
            item[f"{scenario}_CAPA"] = destination.capacity if destination else ""
            item[f"{scenario}_ECART"] = (
                destination.capacity - event["effect"]
                if destination and destination.capacity is not None and event["effect"] is not None
                else ""
            )
        details.append(item)
    return details


def serialize_conflicts(conflicts: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output = []
    for item in conflicts:
        output.append(
            {
                "code": item["code"],
                "room": item["room"],
                "building": item["building"],
                "date": item["date"],
                "slot1": f"{item['start1']//60:02d}:{item['start1']%60:02d}–{item['end1']//60:02d}:{item['end1']%60:02d}",
                "title1": item["title1"],
                "slot2": f"{item['start2']//60:02d}:{item['start2']%60:02d}–{item['end2']//60:02d}:{item['end2']%60:02d}",
                "title2": item["title2"],
                "overlap": item["overlap"],
            }
        )
    return output


def model_template_from_stream(path: Path, destination: Path) -> int:
    iterator = iter_csv_dicts(path)
    metadata = next(iterator)
    if "NOM_DIP" not in metadata["__headers__"]:
        raise SourceError("Colonne NOM_DIP absente ; modèle non créé.")
    names = {row.get("NOM_DIP", "").strip() for row in iterator if row.get("NOM_DIP", "").strip()}
    with destination.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream, delimiter=";", lineterminator="\n")
        writer.writerow(("NOM_DIP", "EFFECTIF_MANUEL"))
        writer.writerows((name, "") for name in sorted(names, key=str.casefold))
    return len(names)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Analyse l'occupation des salles et simule le report des amphis LET.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--dossier", type=Path, default=Path(__file__).resolve().parent / "BDD")
    parser.add_argument("--planning", help="Chemin vers le CSV de planning.")
    parser.add_argument("--salles", help="Chemin vers EXP_SALLE.")
    parser.add_argument("--promotions", help="Chemin vers EXP_PROMOTION.")
    parser.add_argument(
        "--modele-dashboard",
        type=Path,
        default=Path(__file__).resolve().parent / "BDD" / "dashboard_template.html",
        help="Gabarit historique pour le tableau de bord consolidé.",
    )
    parser.add_argument(
        "--surfaces",
        type=Path,
        default=None,
        help="Classeur SURFACES GENERAL (Ville/Site/Code Bât) pour les filtres ville et site ; BDD/SURFACES*.xlsx par défaut.",
    )
    parser.add_argument(
        "--fermetures",
        default=DEFAULT_CLOSURE_PERIODS,
        help="Périodes de fermeture AAAA-MM-JJ:AAAA-MM-JJ (vacances universitaires) séparées par des virgules ; chaîne vide pour aucune.",
    )
    parser.add_argument(
        "--examens",
        default=DEFAULT_EXAM_PERIODS,
        help="Périodes d'examen AAAA-MM-JJ:AAAA-MM-JJ séparées par des virgules ; chaîne vide pour aucune.",
    )
    parser.add_argument(
        "--effectifs-estimes",
        action="store_true",
        help="Pour les séances des amphis LET sans EFFCALCU et de la forme « <Promotion>groupe », retient l'effectif de la "
        "promotion parent d'EXP_PROMOTION comme majorant (estimation marquée comme telle).",
    )
    parser.add_argument("--sortie", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--debut", help="Date ISO YYYY-MM-DD incluse.")
    parser.add_argument("--fin", help="Date ISO YYYY-MM-DD incluse.")
    parser.add_argument("--exclure-semaines", help="Semaines ISO à exclure, ex. 43,52,1.")
    parser.add_argument("--modele-effectifs", action="store_true", help="Écrit les NOM_DIP distincts à renseigner puis s'arrête.")
    parser.add_argument("--self-test", action="store_true", help="Exécute les tests synthétiques du lecteur et des scénarios.")
    return parser


def validate_scenario(name: str, result: dict[str, Any], sources: list[dict[str, Any]], recipients: list[Room], strict_no_overlap: bool) -> None:
    if result["source_minutes"] != result["accounted_minutes"]:
        raise AnalysisError(f"{name} : heures-séances source non conservées.")
    if len(result["statuses"]) != len(sources):
        raise AnalysisError(f"{name} : des séances sources ont disparu du bilan.")
    if any(status["status"] not in {"place", "place_avec_chevauchement", "sans_solution", "non_couvert", "effectif_inconnu"} for status in result["statuses"].values()):
        raise AnalysisError(f"{name} : statut de séance non reconnu.")
    if strict_no_overlap and result["conflicts"]:
        sample = result["conflicts"][0]
        raise AnalysisError(
            f"{name} : {len(result['conflicts'])} chevauchement(s) sur les locaux d'accueil, premier cas {sample['room']} le {sample['date']}."
        )
    recipient_codes = {room.code: room for room in recipients}
    for event in result["placed"]:
        room = recipient_codes[event["room_code"]]
        if event["effect"] is None or room.capacity is None or event["effect"] > room.capacity:
            raise AnalysisError(f"{name} : capacité insuffisante pour la séance {event['id']}.")


def render_saturation(room_rows: list[dict[str, Any]], week_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output = []
    for room in room_rows:
        for week in room["weekly"]:
            if week["rate"] > 100:
                output.append({"room": room["name"], "building": room["building"], **week})
    return sorted(output, key=lambda item: item["rate"], reverse=True)


def room_week_data(events: list[dict[str, Any]], active_dates: list[dt.date]) -> dict[str, list[dict[str, Any]]]:
    week_days: Counter[tuple[int, int]] = Counter()
    for date in active_dates:
        iso = date.isocalendar()
        week_days[(iso.year, iso.week)] += 1
    minutes: Counter[tuple[str, int, int]] = Counter()
    for event in events:
        duration = active_minutes(event)
        iso = event["date"].isocalendar()
        minutes[(event["room_code"], iso.year, iso.week)] += duration
    output: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for (code, year, week), value in minutes.items():
        output[code].append(
            {
                "year": year,
                "week": week,
                "hours": value / 60,
                "days": week_days[(year, week)],
                "rate": 100 * value / max(1, week_days[(year, week)] * 600),
                "semester": "S1" if week >= 27 else "S2",
            }
        )
    for values in output.values():
        values.sort(key=lambda item: (item["year"], item["week"]))
    return output


def enrich_room_rows(room_rows: list[dict[str, Any]], events: list[dict[str, Any]], active_dates: list[dt.date]) -> list[dict[str, Any]]:
    by_code = room_week_data(events, active_dates)
    for room in room_rows:
        room["weekly"] = by_code.get(room["code"], [])
    return room_rows


def run_self_test() -> None:
    assert parse_decimal("1,50") == 1.5
    assert parse_period_weeks("[52..1]") == [52, 53, 1]
    wrap = dates_from_period("[52..1]", dt.date(2025, 12, 22))
    assert wrap == [dt.date(2025, 12, 22), dt.date(2025, 12, 29)]
    fixture = "NOM_DIP;UID;UID;TYPE;PERIODE;DDEBUT;JOUR;HDEBUT;HFIN;DUREE;NOM_SAL;CODE_SAL;EFFCALCU;LIBELLE_MAT\n"
    fixture += "TD test;u1;u2;TD;[36..37];01/09/2025;lundi;08:00;09:30;1,50;Salle test (DEG-00);R1;40;Test\n"
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp) / "planning.csv"
        path.write_text(fixture, encoding="utf-8-sig")
        rows = iter_csv_dicts(path)
        meta = next(rows)
        row = next(rows)
        assert meta["__headers__"][2] == "UID__2"
        assert row["UID"] == "u1" and row["UID__2"] == "u2"
        assert dates_from_period(row["PERIODE"], dt.date(2025, 9, 1)) == [dt.date(2025, 9, 1), dt.date(2025, 9, 8)]
        rows.close()
        fixture_room = Room("R1", "Salle test (DEG-00)", 50, "DEG", "grande_salle")
        loaded = load_schedule(path, {"R1": fixture_room}, {"TD test"}, None, None, set())
        assert len(loaded["events"]) == 2
        manual = independent_manual_check(
            path, loaded["events"], {"R1"}, dt.date(2025, 9, 1), dt.date(2025, 9, 8), set(), loaded["holidays"]
        )
        assert manual[0]["rawEvents"] == manual[0]["engineEvents"] == 2
        assert manual[0]["rawHours"] == manual[0]["engineHours"] == 3.0
        test_rooms, test_weeks = calculate_room_stats(loaded["events"], {"R1": fixture_room}, loaded["active_dates"])
        test_rooms = enrich_room_rows(test_rooms, loaded["events"], loaded["active_dates"])
        placeholder_names = (
            "DATE_CREATION_ISO SCENARIO NB_SALLES PERIODE_LABEL DATE_CREATION NOTE_CHEVAUCHEMENTS DATE_EXTRACTION "
            "NOTE_REFERENCE ANNEE_S1 ANNEE_S2 JOURS_S1 JOURS_S2 JOURS_TOTAL HEURES_BASE RAWDATA HEATMAPBAT "
            "HEATMAPSALLE HEATMAPSALLES1 HEATMAPSALLES2 WEEKLYDATA RAWCONFLICTS WEEKLYSALLEDATA"
        ).split()
        template_path = Path(temp) / "dashboard_template.html"
        template_path.write_text("<!doctype html><title>{{SCENARIO}}</title>" + "|".join(f"{{{{{name}}}}}" for name in placeholder_names), encoding="utf-8")
        template_page = legacy_template_dashboard(
            template_path, "Test", "2025-09-01T12:00:00+02:00", test_rooms, test_weeks,
            loaded["events"], [], loaded["active_dates"]
        )
        assert "{{" not in template_page and "Test" in template_page
        bundle = combined_template_dashboard(
            "2025-09-01T12:00:00+02:00", "test", {"reference": template_page},
            "06_test.csv", "07_test.csv"
        )
        assert "Référence" in bundle and "07_test.csv" in bundle
    room_a = Room("A", "Amphi A (DEG-00)", 100, "DEG", "amphi")
    room_b = Room("B", "Amphi B (DEG-00)", 200, "DEG", "amphi")
    date = dt.date(2025, 9, 1)
    event = {"id": "T1", "date": date, "start": 540, "end": 600, "duration": 60, "room_code": "LET", "room_name": "Amphi LET", "building": "LET", "capacity": 50, "room_kind": "amphi", "effect": 80, "effect_issue": "", "subject": "Test", "type": "CM", "diplomas": ["TD test"]}
    baseline = []
    h1 = scenario_result("H1", baseline, [event], [room_a, room_b], [date], False)
    assert h1["statuses"]["T1"]["destination_code"] == "A"
    h2 = scenario_result("H2", baseline, [event], [room_a, room_b], [date], True)
    assert h2["statuses"]["T1"]["status"] == "place"
    assert h2["accounted_minutes"] == event["duration"]
    print("Tests synthétiques : CSV à UID répété, plages ISO à cheval sur l'année, effectif/capacité, H1 et H2 : OK.")


# ---------------------------------------------------------------------------
# Indicateurs complémentaires : classements, remplissage, examens, avant/après
# ---------------------------------------------------------------------------

DEFAULT_EXAM_PERIODS = "2026-01-05:2026-01-17,2026-05-18:2026-06-30"
PERIMETER_BUILDINGS = {"LET", "DEG"}
PERIMETER_KINDS = {"amphi", "grande_salle"}


# Fermetures de l'université 2025-2026 (jours ouvrés sans cours) ; les jours fériés nationaux sont déjà retranchés par ailleurs.
# Une vacance « du vendredi soir au lundi matin » ne retire que les jours ouvrés strictement entre ces deux dates.
DEFAULT_CLOSURE_PERIODS = (
    "2025-10-27:2025-10-31,"  # Toussaint
    "2025-12-22:2026-01-02,"  # Noël
    "2026-02-16:2026-02-20,"  # Hiver
    "2026-04-06:2026-04-17"  # Printemps
)


def closed_dates_from(text: str) -> set[dt.date]:
    """Jours de fermeture AAAA-MM-JJ:AAAA-MM-JJ (bornes incluses), séparés par des virgules."""
    dates: set[dt.date] = set()
    for first, last in parse_exam_periods(text):
        dates.update(first + dt.timedelta(days=offset) for offset in range((last - first).days + 1))
    return dates


def parse_exam_periods(text: str) -> list[tuple[dt.date, dt.date]]:
    """Lit « 2026-01-05:2026-01-17,2026-05-18:2026-06-30 » (bornes incluses)."""
    periods: list[tuple[dt.date, dt.date]] = []
    for chunk in (text or "").split(","):
        if not chunk.strip():
            continue
        left, _, right = chunk.partition(":")
        first, last = dt.date.fromisoformat(left.strip()), dt.date.fromisoformat((right or left).strip())
        if first > last:
            raise ValueError(f"Période d'examen inversée : {chunk!r}.")
        periods.append((first, last))
    return periods


def is_exam_date(date: dt.date, periods: list[tuple[dt.date, dt.date]]) -> bool:
    return any(first <= date <= last for first, last in periods)


def perimeter_rooms(rooms: dict[str, Room]) -> list[Room]:
    """Amphis et grandes salles des bâtiments LET/DEG avec une capacité connue."""
    return sorted(
        (r for r in rooms.values() if r.building in PERIMETER_BUILDINGS and r.kind in PERIMETER_KINDS and (r.capacity or 0) > 0),
        key=lambda r: (r.building, r.name.casefold(), r.code),
    )


def occupancy_by_room(
    events: list[dict[str, Any]], rooms: list[Room], dates: list[dt.date]
) -> dict[str, dict[str, float]]:
    """Heures et taux (plage 08:00–18:00) par salle sur l'ensemble de jours donné."""
    date_set = set(dates)
    minutes: Counter[str] = Counter()
    for event in events:
        if event["date"] in date_set:
            minutes[event["room_code"]] += clipped_minutes(event["start"], event["end"])
    denominator = len(date_set) * 600
    return {
        room.code: {
            "hours": minutes[room.code] / 60,
            "rate": 100 * minutes[room.code] / denominator if denominator else 0.0,
        }
        for room in rooms
    }


def fill_by_room(events: list[dict[str, Any]], rooms: list[Room], dates: set[dt.date] | None) -> dict[str, dict[str, Any]]:
    """Taux de remplissage moyen (effectif / capacité) et écart moyen par salle."""
    codes = {room.code for room in rooms}
    ratios: dict[str, list[float]] = defaultdict(list)
    gaps: dict[str, list[int]] = defaultdict(list)
    for event in events:
        if event["room_code"] not in codes or event["effect"] is None or not event["capacity"]:
            continue
        if dates is not None and event["date"] not in dates:
            continue
        ratios[event["room_code"]].append(event["effect"] / event["capacity"])
        gaps[event["room_code"]].append(event["capacity"] - event["effect"])
    return {
        code: {
            "sessions": len(ratios[code]),
            "fill": 100 * sum(ratios[code]) / len(ratios[code]),
            "gap": sum(gaps[code]) / len(gaps[code]),
        }
        for code in ratios
    }


def session_gap_rows(
    events: list[dict[str, Any]], rooms: list[Room], dates: set[dt.date] | None
) -> list[dict[str, Any]]:
    """Séances classées par écart décroissant entre capacité et effectif."""
    codes = {room.code for room in rooms}
    rows = [
        {
            "date": event["date"].isoformat(),
            "room": event["room_name"],
            "code": event["room_code"],
            "building": event["building"],
            "type": event["type"],
            "subject": event["subject"],
            "diplomas": " | ".join(event["diplomas"]),
            "effect": event["effect"],
            "capacity": event["capacity"],
            "gap": event["capacity"] - event["effect"],
            "fill": round(100 * event["effect"] / event["capacity"], 1),
        }
        for event in events
        if event["room_code"] in codes
        and event["effect"] is not None
        and event["capacity"]
        and (dates is None or event["date"] in dates)
    ]
    return sorted(rows, key=lambda row: (-row["gap"], row["date"], row["room"]))


def top_flop_rooms(
    occupancy: dict[str, dict[str, float]], rooms: list[Room], count: int = 5
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    by_code = {room.code: room for room in rooms}
    ranked = sorted(
        (
            {"code": code, "room": by_code[code].name, "building": by_code[code].building,
             "capacity": by_code[code].capacity, "hours": values["hours"], "rate": values["rate"]}
            for code, values in occupancy.items()
        ),
        key=lambda row: (-row["rate"], row["room"].casefold()),
    )
    return ranked[:count], ranked[-count:][::-1]


def conflict_breakdown(
    conflicts_by_scenario: dict[str, list[dict[str, Any]]], periods: list[tuple[dt.date, dt.date]]
) -> list[dict[str, Any]]:
    """Chevauchements par scénario, salle, semaine ISO et période (examens / hors examens)."""
    counts: Counter[tuple[str, str, str, str, str]] = Counter()
    for scenario, conflicts in conflicts_by_scenario.items():
        for item in conflicts:
            date = dt.date.fromisoformat(item["date"])
            iso = date.isocalendar()
            period = "examens" if is_exam_date(date, periods) else "hors examens"
            counts[(scenario, item["code"], item["room"], f"{iso.year}-S{iso.week:02d}", period)] += 1
    return [
        {"scenario": k[0], "code": k[1], "room": k[2], "week": k[3], "period": k[4], "conflicts": n}
        for k, n in sorted(counts.items())
    ]


def final_status(source: dict[str, Any], status: dict[str, Any] | None) -> str:
    """inchangée / reportée / lissée / non résolue pour une séance source."""
    if status is None:
        return "inchangée"
    if status["status"] == "place":
        moved = status["date"] != source["date"] or status["start"] != source["start"]
        return "lissée" if moved else "reportée"
    if status["status"] == "place_avec_chevauchement":
        return "reportée (chevauchement)"
    return "non résolue"


def anomaly_rows(schedule: dict[str, Any], events: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Journal d'anomalies : tout ce qui n'a pu être exploité tel quel est listé, rien n'est inventé."""
    rows: list[dict[str, Any]] = [
        {"type": "salle absente du catalogue EXP_SALLE", "valeur": code, "nombre": n}
        for code, n in sorted(schedule["unknown_codes"].items())
    ]
    rows.append({"type": "code salle préfixé <> normalisé (nom concordant avec EXP_SALLE)", "valeur": "", "nombre": schedule["prefixed_normalized"]})
    rows += [
        {"type": "code salle préfixé <> rejeté (nom différent ou code inconnu)", "valeur": code, "nombre": n}
        for code, n in sorted(schedule["prefixed_rejected"].items())
    ]
    rows += [
        {"type": "promotion non reconnue dans EXP_PROMOTION", "valeur": name, "nombre": 1}
        for name in schedule["unrecognized_promotions"]
    ]
    missing = Counter((event["effect_issue"], event["subject"]) for event in events if event["effect"] is None)
    rows += [
        {"type": f"effectif inexploitable ({issue})", "valeur": subject, "nombre": n}
        for (issue, subject), n in sorted(missing.items())
    ]
    over = Counter(
        (event["room_name"], event["subject"])
        for event in events
        if event["effect"] is not None and event["capacity"] and event["effect"] > event["capacity"]
    )
    rows += [
        {"type": "effectif EFFCALCU supérieur à la capacité de la salle", "valeur": f"{room} · {subject}", "nombre": n}
        for (room, subject), n in sorted(over.items())
    ]
    for key, label in (
        ("missing_room_code_count", "ligne avec salle nommée mais sans CODE_SAL"),
        ("duration_mismatch", "DUREE incohérente avec HDEBUT/HFIN"),
        ("at_room_count", "ligne multi-salles (@)"),
        ("admin_count", "blocage administratif >= 8 h exclu"),
    ):
        rows.append({"type": label, "valeur": "", "nombre": schedule[key]})
    for row in rows:
        row["explication"] = explain_anomaly(row["type"])
    return rows


def build_indicators(
    events: list[dict[str, Any]],
    rooms: dict[str, Room],
    scenarios: dict[str, dict[str, Any]],
    active_dates: list[dt.date],
    periods: list[tuple[dt.date, dt.date]],
) -> dict[str, Any]:
    perimeter = perimeter_rooms(rooms)
    exam_dates = [date for date in active_dates if is_exam_date(date, periods)]
    exam_set = set(exam_dates)
    streams = {"Référence": events, **{key: result["events"] for key, result in scenarios.items()}}
    occupancy = {name: occupancy_by_room(stream, perimeter, active_dates) for name, stream in streams.items()}
    occupancy_exam = {name: occupancy_by_room(stream, perimeter, exam_dates) for name, stream in streams.items()}
    top, flop = top_flop_rooms(occupancy["Référence"], perimeter)
    top_exam, flop_exam = top_flop_rooms(occupancy_exam["Référence"], perimeter) if exam_dates else ([], [])
    # Séances surdimensionnées (capacité >= 2 x effectif, >= 20 places libres) : déplaçables si un local
    # plus petit est libre à toutes les occurrences de la série (cf. overdimension_rows).
    oversized = overdimension_rows(events, rooms)
    movable = [row for row in oversized if row["replacementCode"]]
    freed_amphi = [row for row in movable if rooms.get(row["code"]) and rooms[row["code"]].kind == "amphi"]
    breakdown = conflict_breakdown({key: result["conflicts"] for key, result in scenarios.items()}, periods)
    return {
        "perimeter": perimeter,
        "examDates": exam_dates,
        "occupancy": occupancy,
        "occupancyExam": occupancy_exam,
        "top": top, "flop": flop, "topExam": top_exam, "flopExam": flop_exam,
        "fill": fill_by_room(events, perimeter, None),
        "fillExam": fill_by_room(events, perimeter, exam_set),
        "gaps": session_gap_rows(events, perimeter, None),
        "gapsExam": session_gap_rows(events, perimeter, exam_set),
        "firstDate": min(active_dates).isoformat(),
        "lastDate": max(active_dates).isoformat(),
        "activeDays": len(active_dates),
        "overCapacity": sum(1 for row in session_gap_rows(events, perimeter, None) if row["gap"] < 0),
        "oversized": len(oversized),
        "movable": len(movable),
        "freedAmphi": len(freed_amphi),
        "freedAmphiHours": sum(row["end"] - row["start"] for row in freed_amphi) / 60,
        "breakdown": breakdown,
    }


def indicator_csvs(directory: Path, stamp: str, indicators: dict[str, Any], rooms: dict[str, Room]) -> list[Path]:
    perimeter: list[Room] = indicators["perimeter"]
    written: list[Path] = []

    def emit(name: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
        path = directory / f"{name}_{stamp}.csv"
        write_csv(path, headers, rows)
        written.append(path)

    names = list(indicators["occupancy"])
    before_after = []
    for room in perimeter:
        row: dict[str, Any] = {"code": room.code, "salle": room.name, "batiment": room.building, "capacite": room.capacity}
        for name in names:
            row[f"heures_{name}"] = round(indicators["occupancy"][name][room.code]["hours"], 2)
            row[f"taux_{name}"] = round(indicators["occupancy"][name][room.code]["rate"], 2)
            if indicators["examDates"]:
                row[f"taux_examens_{name}"] = round(indicators["occupancyExam"][name][room.code]["rate"], 2)
        before_after.append(row)
    emit("09_occupation_avant_apres", list(before_after[0]) if before_after else ["code"], before_after)

    ranking = []
    for scope, top, flop in (
        ("annee", indicators["top"], indicators["flop"]),
        ("examens", indicators["topExam"], indicators["flopExam"]),
    ):
        for label, items in (("plus occupees", top), ("moins occupees", flop)):
            for position, item in enumerate(items, start=1):
                ranking.append(
                    {"perimetre": scope, "classement": label, "rang": position, "salle": item["room"],
                     "batiment": item["building"], "capacite": item["capacity"],
                     "heures": round(item["hours"], 2), "taux": round(item["rate"], 2)}
                )
    emit("10_classement_occupation", ["perimetre", "classement", "rang", "salle", "batiment", "capacite", "heures", "taux"], ranking)

    fill_rows = []
    for scope, data in (("annee", indicators["fill"]), ("examens", indicators["fillExam"])):
        for code, values in sorted(data.items(), key=lambda kv: kv[1]["fill"]):
            fill_rows.append(
                {"perimetre": scope, "salle": rooms[code].name, "batiment": rooms[code].building,
                 "capacite": rooms[code].capacity, "seances": values["sessions"],
                 "remplissage_moyen_pct": round(values["fill"], 1), "ecart_moyen_places": round(values["gap"], 1)}
            )
    emit("11_remplissage_par_salle", ["perimetre", "salle", "batiment", "capacite", "seances", "remplissage_moyen_pct", "ecart_moyen_places"], fill_rows)

    gap_headers = ["date", "room", "code", "building", "type", "subject", "diplomas", "effect", "capacity", "gap", "fill"]
    emit("12_ecarts_effectif_capacite", gap_headers, indicators["gaps"])
    emit("12_ecarts_effectif_capacite_examens", gap_headers, indicators["gapsExam"])
    emit("13_chevauchements_par_salle_semaine", ["scenario", "code", "room", "week", "period", "conflicts"], indicators["breakdown"])
    return written


def indicators_html(
    created_at: str,
    indicators: dict[str, Any],
    summaries: dict[str, dict[str, Any]],
    periods: list[tuple[dt.date, dt.date]],
    schedule: dict[str, Any],
    anomalies: list[dict[str, Any]],
) -> str:
    perimeter: list[Room] = indicators["perimeter"]
    deg_amphis = [room for room in perimeter if room.building == "DEG" and room.kind == "amphi"]

    def mean_rate(name: str) -> float:
        values = [indicators["occupancy"][name][room.code]["rate"] for room in deg_amphis]
        return sum(values) / len(values) if values else 0.0

    def kpi(label: str, value: str) -> str:
        return f'<div class="kpi"><span>{safe_text(label)}</span><strong>{safe_text(value)}</strong></div>'

    def bars(items: list[dict[str, Any]]) -> str:
        return "".join(
            f'<div class="barrow"><span title="{safe_text(i["room"])}">{safe_text(i["room"])}</span>'
            f'<span class="track"><span class="fill {"crit" if i["rate"] > 100 else ""}" style="display:block;width:{min(100, i["rate"]):.1f}%"></span></span>'
            f'<b>{i["rate"]:.1f} %</b></div>'
            for i in items
        ) or "<p class='muted'>Aucune donnée.</p>"

    def table(headers: tuple[str, ...], rows: list[tuple[Any, ...]]) -> str:
        head = "".join(f"<th>{safe_text(h)}</th>" for h in headers)
        body = "".join("<tr>" + "".join(f"<td>{safe_text(c)}</td>" for c in row) + "</tr>" for row in rows)
        return f'<div class="tablewrap"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'

    by_code = {room.code: room for room in perimeter}

    def fill_table(data: dict[str, dict[str, Any]]) -> str:
        rows = [
            (by_code[c].name, by_code[c].building, by_code[c].capacity, v["sessions"], f"{v['fill']:.1f} %", f"{v['gap']:.0f}")
            for c, v in sorted(data.items(), key=lambda kv: kv[1]["fill"])
        ]
        return table(("Salle", "Bât.", "Places", "Séances", "Remplissage", "Écart moyen"), rows)

    def gap_table(rows: list[dict[str, Any]]) -> str:
        return table(
            ("Date", "Salle", "Type", "Matière", "Effectif", "Capacité", "Écart"),
            [(r["date"], r["room"], r["type"], r["subject"], r["effect"], r["capacity"], r["gap"]) for r in rows[:15]],
        )

    resolution = []
    for raw, smooth in (("H1", "H2"), ("H3a", "H3b")):
        before, after = summaries[raw]["conflicts"], summaries[smooth]["conflicts"]
        resolution.append(
            (f"{raw} → {smooth}", before, after, max(0, before - after),
             summaries[smooth]["noSolution"], summaries[smooth]["uncovered"], summaries[smooth]["unknownEffect"])
        )
    by_period: Counter[tuple[str, str]] = Counter()
    for item in indicators["breakdown"]:
        by_period[(item["scenario"], item["period"])] += item["conflicts"]
    exam_text = ", ".join(f"{a.strftime('%d/%m/%Y')}–{b.strftime('%d/%m/%Y')}" for a, b in periods) or "aucune"
    notes = [
        f"Périmètre des classements : amphis et grandes salles des bâtiments {', '.join(sorted(PERIMETER_BUILDINGS))} ayant une capacité renseignée ({len(perimeter)} locaux).",
        "Heures théoriques : jours pédagogiques observés dans l'export × 10 h (08:00–18:00), du lundi au vendredi.",
        f"Périodes d'examen utilisées : {exam_text} (option --examens). Seul le calendrier détermine « examen » ; le type de séance n'est pas utilisé.",
        "Lissage : durée inchangée, capacité respectée, plage 08:00–18:00, aucune promotion (NOM_DIP) à deux endroits au même moment. Les enseignants ne figurent pas dans l'export : leur disponibilité n'est pas contrôlée.",
        f"Séances sans effectif exploitable : {schedule['missing_effect']} ; elles sont exclues des taux de remplissage et listées dans le journal d'anomalies ({len(anomalies)} lignes).",
        f"Période réellement couverte par l'export : {indicators['firstDate']} au {indicators['lastDate']} ({indicators['activeDays']} jours pédagogiques) ; les périodes d'examen hors de cette fenêtre (ex. mai–juin) ne contiennent aucune donnée.",
        f"{indicators['overCapacity']} séances du périmètre ont un effectif EFFCALCU supérieur à la capacité de la salle : elles sont conservées telles quelles (remplissage > 100 %) et listées dans le journal d'anomalies.",
        f"{schedule['blank_counts'].get('CODE_SAL', 0)} lignes sur {schedule['source_rows']} de l'export n'ont aucun CODE_SAL (séances sans salle attribuée dans l'export) : elles comptent pour les jours ouverts mais pas dans l'occupation d'une salle. Si certaines se déroulent en réalité dans les amphis LET ou DEG, leur occupation est sous-estimée.",
        f"{schedule['prefixed_normalized']} lignes avaient un CODE_SAL préfixé « <> » (salles de Bayonne, Anglet, Mont-de-Marsan) : le préfixe est retiré quand le nom correspond exactement à EXP_SALLE. Aucune de ces lignes ne concerne LET ou DEG.",
        "Un taux supérieur à 100 % signale des chevauchements conservés, pas une capacité physique supérieure.",
        "Surdimensionnement : capacité ≥ 2 × effectif et ≥ 20 places inutilisées ; une séance est « déplaçable » si un local assez grand est libre à toutes les occurrences de sa série. Estimation théorique, non validée par les services.",
    ]
    kpis = "".join(
        (
            kpi("Chevauchements bruts (H1)", str(summaries["H1"]["conflicts"])),
            kpi("Restants après lissage (H2)", str(summaries["H2"]["conflicts"])),
            kpi("Bruts grandes salles (H3a)", str(summaries["H3a"]["conflicts"])),
            kpi("Restants après lissage (H3b)", str(summaries["H3b"]["conflicts"])),
            kpi("Occupation amphis DEG avant", f"{mean_rate('Référence'):.1f} %"),
            kpi("Amphis DEG après H1", f"{mean_rate('H1'):.1f} %"),
            kpi("Amphis DEG après H2", f"{mean_rate('H2'):.1f} %"),
            kpi("Amphis DEG après H3b", f"{mean_rate('H3b'):.1f} %"),
        )
    )
    body = f"""<main><header><div><p class="eyebrow">UPPA · Direction du Patrimoine · SSH</p><h1>Indicateurs clés</h1><p class="muted">Créé le {safe_text(created_at)}</p></div></header>
<section class="card"><h2>Chiffres clés</h2><div class="kpis">{kpis}</div></section>
<section class="card"><h2>Chevauchements : résolus et non résolus</h2>{table(("Passage", "Chevauchements bruts", "Restants", "Résolus", "Séances sans solution", "Non couvertes (capacité)", "Effectif inconnu"), resolution)}<h3>Par période</h3>{table(("Scénario", "Période", "Chevauchements"), [(s, p, n) for (s, p), n in sorted(by_period.items())])}<p class="muted">Détail par salle et par semaine : 13_chevauchements_par_salle_semaine.csv.</p></section>
<section class="card"><h2>Salles les plus et les moins occupées (année)</h2><div class="grid2"><div><h3>Top 5</h3>{bars(indicators["top"])}</div><div><h3>Flop 5</h3>{bars(indicators["flop"])}</div></div></section>
<section class="card"><h2>Focus périodes d'examen</h2><p class="muted">{safe_text(exam_text)} · {len(indicators["examDates"])} jours pédagogiques concernés.</p><div class="grid2"><div><h3>Top 5 examens</h3>{bars(indicators["topExam"])}</div><div><h3>Flop 5 examens</h3>{bars(indicators["flopExam"])}</div></div><h3>Plus forts écarts effectif / capacité en examens</h3>{gap_table(indicators["gapsExam"])}<h3>Remplissage moyen par salle (examens)</h3>{fill_table(indicators["fillExam"])}</section>
<section class="card"><h2>Écart effectif / capacité (année)</h2><h3>15 plus forts écarts</h3>{gap_table(indicators["gaps"])}<h3>Remplissage moyen par salle</h3>{fill_table(indicators["fill"])}</section>
<section class="card"><h2>Séances mal dimensionnées déplaçables</h2><p>{indicators["oversized"]} occurrences surdimensionnées ; {indicators["movable"]} ont un local plus adapté libre sur toute leur série ; {indicators["freedAmphi"]} occurrences en amphithéâtre seraient libérées ({indicators["freedAmphiHours"]:.1f} h).</p></section>
<section class="card"><h2>Explication de chaque cas</h2>{"".join(f"<h3>{safe_text(CASE_EXPLANATIONS[k]['title'])}</h3><p>{safe_text(CASE_EXPLANATIONS[k]['what'])} {safe_text(CASE_EXPLANATIONS[k]['how'])}</p>" for k in ("reference", "H1", "H2", "H3a", "H3b"))}<h3>Statuts des séances</h3><ul>{"".join(f"<li><code>{safe_text(k)}</code> : {safe_text(v)}</li>" for k, v in STATUS_EXPLANATIONS.items())}</ul></section>
<section class="card"><h2>Hypothèses et limites</h2><ul>{"".join(f"<li>{safe_text(n)}</li>" for n in notes)}</ul></section></main>"""
    return html_document("Indicateurs occupation UPPA", created_at, body, BASE_CSS)


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.self_test:
        run_self_test()
        return 0
    try:
        start = dt.date.fromisoformat(args.debut) if args.debut else None
        end = dt.date.fromisoformat(args.fin) if args.fin else None
        if start and end and start > end:
            parser.error("--debut doit être antérieur ou égal à --fin.")
    except ValueError:
        parser.error("--debut et --fin doivent respecter le format YYYY-MM-DD.")
    try:
        exam_periods = parse_exam_periods(args.examens)
        closed_dates = closed_dates_from(args.fermetures)
    except ValueError as error:
        parser.error(f"--examens ou --fermetures invalide : {error}")
    try:
        excluded_weeks = {int(value.strip()) for value in args.exclure_semaines.split(",") if value.strip()} if args.exclure_semaines else set()
        if any(week < 1 or week > 53 for week in excluded_weeks):
            raise ValueError
    except ValueError:
        parser.error("--exclure-semaines attend des numéros ISO entre 1 et 53, séparés par des virgules.")

    directory = args.dossier.expanduser().resolve()
    schedule_path = find_source(directory, args.planning, ("CC*.csv", "*PLANNING*.csv"))
    room_path = find_source(directory, args.salles, ("EXP_SALLE*.csv", "*SALLE*.csv"))
    promotion_path = find_source(directory, args.promotions, ("EXP_PROMOTION*.csv", "*PROMOTION*.csv"))
    if not schedule_path or not room_path or not promotion_path:
        print("Erreur : planning, EXP_SALLE et EXP_PROMOTION doivent être présents ou indiqués explicitement.", file=sys.stderr)
        return 2
    try:
        rooms, room_notes = load_rooms(room_path)
        promotion_iterator = iter_csv_dicts(promotion_path)
        promotion_metadata = next(promotion_iterator)
        if not PROMOTION_REQUIRED.issubset(set(promotion_metadata["__headers__"])):
            raise SourceError(f"Colonnes requises absentes de {promotion_path.name}.")
        promotion_effects: dict[str, set[int]] = defaultdict(set)
        promotion_names: set[str] = set()
        for row in promotion_iterator:
            name = row.get("NOM_DIP", "").strip()
            if name:
                promotion_names.add(name)
                value = parse_effect(str(row.get("EFFCALCU", "")).strip())
                if value is not None:
                    promotion_effects[name].add(value)
        schedule = load_schedule(schedule_path, rooms, promotion_names, start, end, excluded_weeks, closed_dates)
        if schedule["iso_mismatch"] or schedule["day_mismatch"]:
            raise AnalysisError(
                f"Contrôle calendrier échoué : {schedule['iso_mismatch']} PERIODE/DDEBUT incohérentes, {schedule['day_mismatch']} JOUR/DDEBUT incohérents."
            )
        if start is None:
            start = min(schedule["active_dates"])
        if end is None:
            end = max(schedule["active_dates"])
        active_dates = [date for date in schedule["active_dates"] if start <= date <= end]
        schedule["events"] = [event for event in schedule["events"] if start <= event["date"] <= end]
        if not active_dates:
            raise AnalysisError("Aucun jour pédagogique ouvré ne reste après les filtres.")
        events = schedule["events"]
        source_codes = {room.code for room in rooms.values() if room.building == "LET" and room.kind == "amphi"}
        source_events = [event for event in events if event["room_code"] in source_codes]
        schedule["missing_effect"] = sum(event["effect"] is None for event in events)
        estimated_count = 0
        if args.effectifs_estimes:
            # Majorant : pour une séance sans EFFCALCU dont l'unique promotion s'écrit « <Parent>groupe », on retient l'effectif
            # de la promotion parent dans EXP_PROMOTION (borne haute, valable dans 96 % des séances contrôlées). Limité aux
            # séances à reporter : la référence et les classements gardent les seuls effectifs du planning.
            estimated_sources = []
            for event in source_events:
                parent = re.match(r"<([^>]*)>", event["diplomas"][0]) if len(event["diplomas"]) == 1 else None
                values = {v for v in promotion_effects.get(parent.group(1), ())} if parent else set()
                if event["effect"] is None and values:
                    event = dict(
                        event, effect=max(values), effect_estimated=True,
                        effect_issue=f"majorant : effectif de la promotion parent « {parent.group(1)} »",
                    )
                    estimated_count += 1
                estimated_sources.append(event)
            source_events = estimated_sources
        schedule["estimated_sources"] = estimated_count
        print(f"Effectifs estimés (majorant promotion parent) : {estimated_count} séances à reporter.")
        schedule["missing_effect_source"] = sum(event["effect"] is None for event in source_events)
        schedule["unrecognized_source_promotions"] = sorted(
            {diploma for event in source_events for diploma in event["diplomas"] if diploma not in promotion_names},
            key=str.casefold,
        )
        print(f"Séances sans effectif retrouvé : {schedule['missing_effect']} localisées ; {schedule['missing_effect_source']} dans les amphis LET à reporter.")
        if schedule["unrecognized_promotions"]:
            print("Promotions non reconnues (égalité exacte NOM_DIP / EXP_PROMOTION) :")
            for diploma in schedule["unrecognized_promotions"]:
                print(f"- {diploma}")
        else:
            print("Promotions non reconnues : aucune (égalité exacte NOM_DIP / EXP_PROMOTION).")
        print(f"Parmi elles, présentes dans les amphis LET : {len(schedule['unrecognized_source_promotions'])}.")

        if args.modele_effectifs:
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            destination = directory / f"effectifs_a_renseigner_{stamp}.csv"
            count = model_template_from_stream(schedule_path, destination)
            print(f"Modèle créé : {destination} ({count} NOM_DIP distincts, recopiés sans normalisation).")
            return 0

        recipients_h1 = [room for room in rooms.values() if room.building == "DEG" and room.kind == "amphi"]
        recipients_h3 = [
            room
            for room in rooms.values()
            if (room.building == "DEG" and room.kind == "amphi")
            or (room.building in {"LET", "DEG"} and room.kind == "grande_salle")
        ]
        if len(recipients_h1) != 5:
            raise AnalysisError(f"5 amphithéâtres DEG attendus dans EXP_SALLE, {len(recipients_h1)} trouvés.")
        if len([room for room in rooms.values() if room.building == "LET" and room.kind == "amphi"]) != 3:
            raise AnalysisError("Le catalogue ne contient pas exactement trois amphithéâtres LET identifiables.")
        if not recipients_h3:
            raise AnalysisError("Aucun local d'accueil H3 identifié dans LET/DEG.")

        source_ids = {event["id"] for event in source_events}
        baseline_without_source = [event for event in events if event["id"] not in source_ids]
        let_amp_codes = {
            room.code for room in rooms.values() if room.building == "LET" and room.kind == "amphi"
        }
        manual_checks = independent_manual_check(
            schedule_path, events, let_amp_codes, start, end, excluded_weeks, schedule["holidays"]
        )
        h1 = scenario_result("H1", baseline_without_source, source_events, recipients_h1, active_dates, False)
        h2 = scenario_result("H2", baseline_without_source, source_events, recipients_h1, active_dates, True)
        h3a = scenario_result("H3a", baseline_without_source, source_events, recipients_h3, active_dates, False)
        h3b = scenario_result("H3b", baseline_without_source, source_events, recipients_h3, active_dates, True)
        validate_scenario("H1", h1, source_events, recipients_h1, False)
        validate_scenario("H2", h2, source_events, recipients_h1, True)
        validate_scenario("H3a", h3a, source_events, recipients_h3, False)
        validate_scenario("H3b", h3b, source_events, recipients_h3, True)

        run_directory, stamp, created_at = create_run_directory(args.sortie.expanduser().resolve())
        ref_conflicts = find_conflicts(events)
        reference_rows, reference_weeks = calculate_room_stats(events, rooms, active_dates)
        reference_rows = enrich_room_rows(reference_rows, events, active_dates)
        reference_summary = {
            "label": "Occupation de référence 2025-2026",
            "placed": 0,
            "withConflicts": 0,
            "noSolution": 0,
            "uncovered": 0,
            "unknownEffect": schedule["missing_effect"],
            "placedHours": sum(active_minutes(event) for event in events) / 60,
            "sourceHours": sum(active_minutes(event) for event in events) / 60,
            "accountedHours": sum(active_minutes(event) for event in events) / 60,
            "conflicts": len(ref_conflicts),
        }
        scenarios: dict[str, dict[str, Any]] = {"H1": h1, "H2": h2, "H3a": h3a, "H3b": h3b}
        scenario_summaries: dict[str, dict[str, Any]] = {}
        scenario_pages: list[tuple[str, dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]] = []
        for key, result in scenarios.items():
            summary = summarize_scenario(result)
            scenario_summaries[key] = summary
            rows, weeks = calculate_room_stats(result["events"], rooms, active_dates)
            rows = enrich_room_rows(rows, result["events"], active_dates)
            scenario_pages.append((key, result, rows, weeks))

        source_week_rows = room_week_data(events, active_dates)
        saturated = []
        for room in reference_rows:
            for week in room["weekly"]:
                if week["rate"] > 100:
                    saturated.append({"room": room["name"], "building": room["building"], **week})
        saturated.sort(key=lambda row: row["rate"], reverse=True)

        over_rows = overdimension_rows(events, rooms)
        source_total_hours = sum(event["duration"] for event in source_events) / 60
        calendar_info = {"start": start.isoformat(), "end": end.isoformat()}
        summary_data = dict(schedule)
        summary_data["room_notes"] = room_notes
        summary_html_text = summary_html(
            created_at, stamp, summary_data, scenario_summaries, reference_summary["placedHours"],
            saturated, calendar_info, manual_checks
        )
        summary_path = run_directory / f"00_synthese_{stamp}.html"
        summary_path.write_text(summary_html_text, encoding="utf-8")
        ref_path = run_directory / f"01_reference_2025-2026_{stamp}.html"
        ref_path.write_text(
            render_dashboard("Occupation des salles · Référence 2025-2026", "Occupation issue des séances localisées", created_at, reference_rows, reference_weeks, serialize_conflicts(ref_conflicts), reference_summary, len(active_dates), stamp),
            encoding="utf-8",
        )
        scenario_names = {"H1": "02_H1_reel_brut", "H2": "03_H2_lissage", "H3a": "04_H3a_grandes_salles_brut", "H3b": "05_H3b_grandes_salles_lissage"}
        for key, result, rows, weeks in scenario_pages:
            page_path = run_directory / f"{scenario_names[key]}_{stamp}.html"
            page_path.write_text(
                render_dashboard(f"Occupation des salles · {key}", "Report des séances des amphithéâtres LET", created_at, rows, weeks, serialize_conflicts(result["conflicts"]), scenario_summaries[key], len(active_dates), stamp),
                encoding="utf-8",
            )

        over_path = run_directory / f"06_surdimensionnement_{stamp}.html"
        over_html_text = overdimension_html(over_rows, created_at)
        over_path.write_text(over_html_text, encoding="utf-8")
        over_csv = run_directory / f"06_surdimensionnement_{stamp}.csv"
        write_csv(over_csv, ["date", "room", "code", "building", "capacity", "effect", "unused", "type", "subject", "diplomas", "seriesOccurrences", "replacement", "replacementCode"], over_rows)
        details = detail_rows(events, scenarios, rooms)
        detail_csv = run_directory / f"07_detail_seances_{stamp}.csv"
        detail_headers = list(details[0]) if details else ["ID_SEANCE", "DATE", "CODE_SAL"]
        write_csv(detail_csv, detail_headers, details)

        indicators = build_indicators(events, rooms, scenarios, active_dates, exam_periods)
        anomalies = anomaly_rows(schedule, events)
        anomaly_csv = run_directory / f"14_anomalies_{stamp}.csv"
        write_csv(anomaly_csv, ["type", "valeur", "nombre", "explication"], anomalies)
        indicator_files = indicator_csvs(run_directory, stamp, indicators, rooms)
        indicators_page = indicators_html(created_at, indicators, scenario_summaries, exam_periods, schedule, anomalies)
        indicators_path = run_directory / f"15_indicateurs_{stamp}.html"
        indicators_path.write_text(indicators_page, encoding="utf-8")

        template_path = args.modele_dashboard.expanduser().resolve()
        surfaces_path = args.surfaces.expanduser().resolve() if args.surfaces else next(iter(sorted(directory.glob("SURFACES*.xlsx"))), None)
        building_sites = load_building_sites(surfaces_path) if surfaces_path and surfaces_path.is_file() else {}
        if building_sites:
            used = {room["building"] for room in reference_rows}
            unmatched = sorted(code for code in used if code not in building_sites)
            print(f"Villes/sites chargés depuis {surfaces_path.name} ; bâtiments sans ville/site : {', '.join(unmatched) or 'aucun'}.")
        else:
            print("Aucun fichier SURFACES trouvé : filtres ville/site vides.")
        dashboard_pages = {"surdimensionnement": over_html_text}
        if template_path.is_file():
            dashboard_pages["reference"] = legacy_template_dashboard(
                template_path, "— Référence 2025-2026", created_at, reference_rows, reference_weeks,
                events, ref_conflicts, active_dates, building_sites, exam_periods
            )
            for key, result, rows, weeks in scenario_pages:
                dashboard_pages[key] = legacy_template_dashboard(
                    template_path, f"— Scénario {key}", created_at, rows, weeks,
                    result["events"], result["conflicts"], active_dates, building_sites, exam_periods
                )
        else:
            dashboard_pages["reference"] = ref_path.read_text(encoding="utf-8")
            for key in scenarios:
                dashboard_pages[key] = (run_directory / f"{scenario_names[key]}_{stamp}.html").read_text(encoding="utf-8")
        rows_by_key = {key: rows for key, _, rows, _ in scenario_pages}
        stories = {
            key: scenario_story(
                key, result, scenario_summaries[key], source_events,
                recipients_h3 if key in {"H3a", "H3b"} else recipients_h1,
                reference_rows, rows_by_key[key], indicators["occupancy"], ref_conflicts,
            )
            for key, result in scenarios.items()
        }
        amphis_page = amphis_html(
            created_at, rooms, building_sites,
            {"Référence": events, **{key: result["events"] for key, result in scenarios.items()}}, active_dates, exam_periods,
        )
        (run_directory / f"16_amphis_{stamp}.html").write_text(amphis_page, encoding="utf-8")
        dashboard_pages["amphis"] = amphis_page
        unknown_rows = unknown_effect_rows(source_events)
        unknown_csv = run_directory / f"17_cours_sans_effectif_{stamp}.csv"
        write_csv(
            unknown_csv,
            ["ID_SEANCE", "DATE", "JOUR", "HDEBUT", "HFIN", "NOM_SAL", "TYPE", "LIBELLE_MAT", "NOM_DIP", "ORIGINE", "HEURES", "EFFECTIF_A_RENSEIGNER"],
            unknown_rows,
        )
        unknown_page = unknown_effect_html(created_at, unknown_rows, schedule.get("estimated_sources", 0), len(source_events))
        (run_directory / f"17_cours_sans_effectif_{stamp}.html").write_text(unknown_page, encoding="utf-8")
        dashboard_pages["sans_effectif"] = unknown_page
        final_dashboard = run_directory / f"08_dashboard_complet_{stamp}.html"
        final_dashboard.write_text(
            combined_template_dashboard(
                created_at,
                stamp,
                dashboard_pages,
                over_csv.name,
                detail_csv.name,
                stories,
                unknown_csv.name,
            ),
            encoding="utf-8",
        )

        print(f"Dossier de résultats : {run_directory}")
        print(f"08 Tableau de bord consolidé depuis le gabarit : {final_dashboard}")
        print(f"00 Synthèse : {summary_path}")
        print(f"01 Référence : {ref_path}")
        for key in scenarios:
            print(f"{key} : {run_directory / f'{scenario_names[key]}_{stamp}.html'}")
        print(f"06 Surdimensionnement HTML : {over_path}")
        print(f"06 Surdimensionnement CSV : {over_csv}")
        print(f"07 Détail séance par séance CSV : {detail_csv}")
        for path in [*indicator_files, anomaly_csv, indicators_path]:
            print(f"{path.name[:2]} {path.name}")
        print(f"Jours pédagogiques retenus : {len(active_dates)} ; fériés nationaux retranchés ; périodes hors calendrier observé exclues.")
        print(f"Heures source des amphithéâtres LET : {source_total_hours:.2f} h ; conservation vérifiée pour chaque scénario (placées + statuts).")
        if ref_conflicts:
            print(f"Chevauchements dans la référence : {len(ref_conflicts)} ; ils restent visibles et ne sont pas réécrits.")
        print("Contrôles H2/H3b : aucun chevauchement dans les locaux d'accueil ; capacités respectées ; toutes les séances classées.")
        return 0
    except (OSError, csv.Error, SourceError, AnalysisError, ValueError) as error:
        print(f"ARRÊT : {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())