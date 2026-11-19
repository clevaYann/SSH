#!/usr/bin/env python3
"""ETL d'occupation des salles UPPA : mêmes calculs que analyse_occupation_uppa.py, mais seulement les CSV.

Aucune page HTML n'est générée. Le script est autonome (il n'importe pas analyse_occupation_uppa.py) : lecture des
sources, simulation H1a/H1b/H1c/H2/H3b, réaffectation globale et indicateurs y sont recopiés. Les CSV produits sont
identiques à ceux de l'analyse complète (mêmes noms, dossier resultats_<horodatage>). Aucun effectif n'est estimé.

Sources (dossier --dossier, par défaut BDD/) : planning (CC*.csv), catalogue des salles EXP_SALLE (--salles pour
indiquer un autre fichier), EXP_PROMOTION (--promotions) et, pour les villes/sites, SURFACES*.xlsx (--surfaces).

Exemples :
    python etl_occupation_uppa.py
    python etl_occupation_uppa.py --salles "BDD/EXP_SALLE_2025-2026.csv" --sortie resultats
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import re
import sys
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


class SourceError(Exception):
    """Erreur de lecture d'une source tabulaire."""


def find_source(directory: Path, explicit: str | None, patterns: tuple[str, ...]) -> Path | None:
    if explicit:
        path = Path(explicit).expanduser().resolve()
        return path if path.is_file() else None
    candidates: list[Path] = []
    for pattern in patterns:
        candidates.extend(directory.glob(pattern))
    unique = sorted(set(candidates), key=lambda item: item.name.casefold())
    return unique[0] if len(unique) == 1 else None


def parse_decimal(value: str) -> float | None:
    try:
        return float(value.strip().replace(",", "."))
    except (AttributeError, ValueError):
        return None


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
    unlocated: dict[tuple[Any, ...], dict[str, Any]] = {}
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
                unlocated_entry = unlocated.setdefault(
                    (occurrence, start_minute, end_minute, subject, course_type),
                    {"date": occurrence, "start": start_minute, "end": end_minute, "subject": subject, "type": course_type,
                     "diplomas": set(), "effects": set(), "is_exam": is_exam_session(course_type, subject)},
                )
                if diploma:
                    unlocated_entry["diplomas"].add(diploma)
                if effect is not None:
                    unlocated_entry["effects"].add(effect)
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
                    "is_exam": is_exam_session(course_type, subject),
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
        "unlocated": list(unlocated.values()),
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
                            "id1": first["id"], "id2": second["id"],
                            "unc1": first["effect"] is None or bool(first.get("effect_estimated")),
                            "unc2": second["effect"] is None or bool(second.get("effect_estimated")),
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
        "source_active_minutes": sum(active_minutes(source) for source in sources),
        "accounted_active_minutes": sum(active_minutes(event) for event in placed) + sum(
            active_minutes(source) for source in sources
            if status_by_id[source["id"]]["status"] not in {"place", "place_avec_chevauchement"}
        ),
        "conflicts": find_conflicts(final_events, {room.code for room in recipients}),
    }


def active_minutes(event: dict[str, Any]) -> int:
    return clipped_minutes(event["start"], event["end"])


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


PERIMETER_BUILDINGS = {"LET", "DEG"}


PERIMETER_KINDS = {"amphi", "grande_salle"}


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
        "freedAmphiHours": sum(clipped_minutes(row["start"], row["end"]) for row in freed_amphi) / 60,
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


EXAM_TYPE_PATTERN = re.compile(r"exam|contr[oô]le|\bcc\b|rattrapage|soutenance|\boral\b|partiel|tp not|devoir", re.I)


EXAM_SUBJECT_PATTERN = re.compile(r"exam|partiel|devoir|rattrapage|soutenance|\boral\b|\bcc\b|contr[oô]le continu|\bcc[_ ]", re.I)


REALLOC_TYPES = {"cours", "cm", "td", "ctd"}


REALLOC_MIN_GAIN = 20  # une réaffectation doit gagner au moins 20 places de capacité


WEEKDAYS_FR = ("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")


HYPOTHESIS_LABELS = {
    "H1a": "5 amphis DEG, mêmes jours et horaires",
    "H1b": "5 amphis DEG + grandes salles LET/DEG (50-100 places), mêmes jours et horaires",
    "H1c": "5 amphis DEG + autres amphis du campus de Pau, mêmes jours et horaires",
    "H2": "Comparaison : 5 amphis DEG avec lissage",
    "H3b": "Comparaison : 5 amphis DEG + grandes salles avec lissage",
}


def is_exam_session(course_type: str, subject: str) -> bool:
    """Examen ou contrôle (type ou intitulé) : contraintes d'accueil spécifiques, distingué des cours ordinaires."""
    return bool(EXAM_TYPE_PATTERN.search(course_type or "") or EXAM_SUBJECT_PATTERN.search(subject or ""))


def event_is_ssh_pau(event: dict[str, Any], names: set[str]) -> bool:
    """Séance d'une promotion de la structure « SSH - Pau » (nom exact ou groupe « <Promotion>... »)."""
    for diploma in event["diplomas"]:
        if diploma in names:
            return True
        parent = re.match(r"<([^>]*)>", diploma)
        if parent and parent.group(1) in names:
            return True
    return False


def reallocation_pool(rooms: dict[str, Room], building_sites: dict[str, dict[str, str]]) -> dict[str, Room]:
    """Locaux banalisés du campus de Pau (amphis, salles de cours, salles numérotées), toutes capacités."""
    pool: dict[str, Room] = {}
    for room in rooms.values():
        if building_sites.get(room.building, {}).get("site") != "CAMPUS DE PAU" or not room.capacity:
            continue
        if room.kind in {"amphi", "salle_cours", "grande_salle"} or re.match(r"salle\s*\d+\s*(bis)?\s*\(", room.name, re.I):
            pool[room.code] = room
    return pool


def hypothesis_rows(
    scenarios: dict[str, dict[str, Any]], source_events: list[dict[str, Any]], recipients_by_key: dict[str, list[Room]]
) -> list[dict[str, Any]]:
    """Bilan chiffré (séances et heures) de chaque hypothèse de report."""
    by_id = {event["id"]: event for event in source_events}
    rows = []
    for key, result in scenarios.items():
        acc: dict[str, list[float]] = defaultdict(lambda: [0, 0.0])

        def add(name: str, event: dict[str, Any]) -> None:
            acc[name][0] += 1
            acc[name][1] += active_minutes(event) / 60

        for source_id, status in result["statuses"].items():
            event = by_id[source_id]
            kind = status["status"]
            add("total", event)
            add("exam" if event.get("is_exam") else "ordinary", event)
            if kind in {"place", "place_avec_chevauchement"}:
                add("estimated" if status.get("estimated") else "placed", event)
                if kind == "place_avec_chevauchement":
                    add("overlap", event)
            elif kind == "effectif_inconnu":
                add("no_effect", event)
            elif kind == "sans_solution":
                add("no_solution", event)
            elif kind == "non_couvert":
                add("uncovered", event)
        unresolved = [
            acc["no_effect"][0] + acc["estimated"][0] + acc["no_solution"][0] + acc["uncovered"][0],
            acc["no_effect"][1] + acc["estimated"][1] + acc["no_solution"][1] + acc["uncovered"][1],
        ]
        row: dict[str, Any] = {
            "key": key, "label": HYPOTHESIS_LABELS.get(key, key), "rooms": len(recipients_by_key[key]),
            "conflicts": len(result["conflicts"]),
        }
        for name in ("total", "ordinary", "exam", "placed", "estimated", "overlap", "no_effect", "no_solution", "uncovered"):
            row[f"{name}_n"], row[f"{name}_h"] = acc[name][0], round(acc[name][1], 2)
        row["unresolved_n"], row["unresolved_h"] = unresolved[0], round(unresolved[1], 2)
        row["top_conflict_rooms"] = Counter(item["room"] for item in result["conflicts"]).most_common(5)
        rows.append(row)
    return rows


def simulate_global_reallocation(
    events: list[dict[str, Any]],
    rooms: dict[str, Room],
    pool: dict[str, Room],
    ssh_names: set[str],
    active_dates: list[dt.date],
    resolve_conflicts: bool = True,
) -> dict[str, Any]:
    """Réaffectation globale : chaque série de cours/TD SSH Pau mal dimensionnée va vers le plus petit local libre à toutes ses
    dates (mêmes jours et horaires). Les séries sont traitées par gaspillage décroissant et l'occupation est mise à jour après
    chaque décision : l'ensemble des propositions est donc compatible."""
    scope = [e for e in events if e["room_code"] in pool and event_is_ssh_pau(e, ssh_names)]
    candidates = [e for e in scope if not e.get("is_exam") and e["type"].strip().casefold() in REALLOC_TYPES]
    candidate_ids = {e["id"] for e in candidates}
    exam_events = [e for e in scope if e.get("is_exam")]
    other_events = [e for e in scope if e["id"] not in candidate_ids and not e.get("is_exam")]

    occupied: dict[tuple[str, dt.date], list[list[Any]]] = defaultdict(list)
    for event in events:
        occupied[(event["room_code"], event["date"])].append([event["start"], event["end"], event["id"]])

    def free(code: str, date: dt.date, start: int, end: int) -> bool:
        return all(not (start < other_end and other_start < end) for other_start, other_end, _ in occupied.get((code, date), []))

    pool_list = sorted(pool.values(), key=lambda r: (r.capacity or 0, r.code))
    moved: dict[str, Room] = {}
    conflict_moved: dict[str, Room] = {}
    conflict_pairs = 0
    if resolve_conflicts:
        # Étape 1 : relogement des séances en double réservation dans un local libre de capacité suffisante (mêmes jour et horaire).
        scope_ids = {e["id"] for e in scope}
        pairs = []
        for (code, _date), items in room_events(events).items():
            if code not in pool:
                continue
            for index, first in enumerate(items):
                for second in items[index + 1:]:
                    if second["start"] >= first["end"]:
                        break
                    if overlap(first, second):
                        pairs.append((first, second))
        conflict_pairs = len(pairs)
        for first, second in pairs:
            if first["id"] in moved or second["id"] in moved:
                continue
            movable = [e for e in (first, second) if e["id"] in scope_ids and not e.get("is_exam") and e["effect"] is not None]
            for event in sorted(movable, key=lambda e: ("@" not in e["id"], e["effect"])):
                current = rooms[event["room_code"]]
                target = next(
                    (r for r in sorted(pool_list, key=lambda r: (r.capacity or 0, r.building != current.building, r.code))
                     if r.code != current.code and (r.capacity or 0) >= event["effect"] and free(r.code, event["date"], event["start"], event["end"])),
                    None,
                )
                if target is None:
                    continue
                for slot in occupied[(current.code, event["date"])]:
                    if slot[2] == event["id"]:
                        occupied[(current.code, event["date"])].remove(slot)
                        break
                occupied[(target.code, event["date"])].append([event["start"], event["end"], event["id"]])
                moved[event["id"]] = target
                conflict_moved[event["id"]] = target
                break

    series: dict[tuple[Any, ...], list[dict[str, Any]]] = defaultdict(list)
    for event in candidates:
        if event["id"] in conflict_moved:
            continue
        series[(event["room_code"], event["subject"], event["type"], tuple(event["diplomas"]), event["date"].weekday(), event["start"], event["end"])].append(event)

    series_rows: list[dict[str, Any]] = []
    prepared = []
    for key, items in series.items():
        efforts = [e["effect"] for e in items]
        current = rooms[key[0]]
        known = all(value is not None for value in efforts)
        prepared.append((key, items, current, known, max(efforts) if known else None))
    prepared.sort(key=lambda p: (-((p[2].capacity or 0) - (p[4] or 0)) if p[3] else 0, p[0][0], p[0][1]))

    for key, items, current, known, max_effect in prepared:
        hours = sum(active_minutes(e) for e in items) / 60
        row = {
            "room": current.name, "code": current.code, "building": current.building, "kind": current.kind, "capacity": current.capacity,
            "max_effect": max_effect, "mean_effect": round(sum(e["effect"] for e in items) / len(items), 1) if known else None,
            "weekday": WEEKDAYS_FR[key[4]], "start": key[5], "end": key[6], "subject": key[1], "type": key[2],
            "estimated": any(e.get("effect_estimated") for e in items),
            "diplomas": " | ".join(key[3]), "sessions": len(items), "hours": round(hours, 2),
            "target": "", "target_code": "", "target_capacity": None, "target_building": "", "target_kind": "", "freed_hours": 0.0, "status": "",
        }
        if not known:
            row["status"] = "effectif inconnu : non réaffectable"
        elif max_effect > (current.capacity or 0):
            row["status"] = (
                "effectif estimé supérieur à la capacité : non réaffecté" if row["estimated"] else "effectif supérieur à la capacité : à traiter à part"
            )
        else:
            targets = [
                r for r in pool_list
                if r.code != current.code and max_effect <= (r.capacity or 0) <= (current.capacity or 0) - REALLOC_MIN_GAIN
            ]
            targets.sort(key=lambda r: (r.capacity or 0, r.building != current.building, r.code))
            if not targets:
                row["status"] = "déjà adapté : aucun local plus petit utile"
            else:
                chosen = next((r for r in targets if all(free(r.code, e["date"], e["start"], e["end"]) for e in items)), None)
                if chosen is None:
                    row["status"] = "non réaffecté : aucun local plus adapté libre à toutes les dates"
                else:
                    for event in items:
                        for slot in occupied[(current.code, event["date"])]:
                            if slot[2] == event["id"]:
                                occupied[(current.code, event["date"])].remove(slot)
                                break
                        occupied[(chosen.code, event["date"])].append([event["start"], event["end"], event["id"]])
                        moved[event["id"]] = chosen
                    row.update(
                        status="réaffecté", target=chosen.name, target_code=chosen.code, target_capacity=chosen.capacity,
                        target_building=chosen.building, target_kind=chosen.kind, freed_hours=round(hours, 2),
                    )
        series_rows.append(row)

    final_events = []
    session_rows = []
    for event in events:
        target = moved.get(event["id"])
        if target is None:
            final_events.append(event)
            continue
        final_events.append(
            dict(event, room_code=target.code, room_name=target.name, building=target.building, capacity=target.capacity,
                 room_kind=target.kind, realloc_from=event["room_code"])
        )
    for event in candidates + [e for e in scope if e["id"] in conflict_moved and e["id"] not in candidate_ids]:
        target = moved.get(event["id"])
        session_rows.append(
            {
                "ID_SEANCE": event["id"], "DATE": event["date"].isoformat(), "JOUR": WEEKDAYS_FR[event["date"].weekday()],
                "HDEBUT": f"{event['start'] // 60:02d}:{event['start'] % 60:02d}", "HFIN": f"{event['end'] // 60:02d}:{event['end'] % 60:02d}",
                "TYPE": event["type"], "LIBELLE_MAT": event["subject"], "NOM_DIP": " | ".join(event["diplomas"]),
                "LOCAL_INITIAL": event["room_name"], "CAPACITE_INITIALE": event["capacity"],
                "EFFECTIF": event["effect"] if event["effect"] is not None else "",
                "LOCAL_PROPOSE": target.name if target else event["room_name"],
                "CAPACITE_PROPOSEE": target.capacity if target else event["capacity"],
                "STATUT": ("déplacée (résolution de conflit)" if event["id"] in conflict_moved else "réaffectée") if target else "inchangée",
                "HEURES": round(active_minutes(event) / 60, 2),
            }
        )

    # Contrôles : aucun conflit créé, capacité respectée, heures conservées
    by_room = room_events(final_events)
    for event in final_events:
        if event["id"] not in moved:
            continue
        if event["effect"] is None or event["effect"] > (event["capacity"] or 0):
            raise AnalysisError(f"Réaffectation : capacité dépassée pour la séance {event['id']}.")
        for other in by_room[(event["room_code"], event["date"])]:
            if other["id"] != event["id"] and overlap(event, other):
                raise AnalysisError(f"Réaffectation : conflit créé dans {event['room_name']} le {event['date']}.")
    if abs(sum(e["duration"] for e in events) - sum(e["duration"] for e in final_events)) > 0:
        raise AnalysisError("Réaffectation : les heures ne sont pas conservées.")

    # Indicateurs avant / après
    pool_rooms = list(pool.values())
    amphis = [r for r in pool_rooms if r.kind == "amphi"]
    before_occ = occupancy_by_room(events, pool_rooms, active_dates)
    after_occ = occupancy_by_room(final_events, pool_rooms, active_dates)

    def weighted_fill(stream: list[dict[str, Any]], ids: set[str]) -> float:
        num = sum((e["effect"] or 0) * active_minutes(e) for e in stream if e["id"] in ids and e["effect"] is not None and e["capacity"])
        den = sum(e["capacity"] * active_minutes(e) for e in stream if e["id"] in ids and e["effect"] is not None and e["capacity"])
        return 100 * num / den if den else 0.0

    known_ids = {e["id"] for e in candidates if e["effect"] is not None}
    fill_before = fill_by_room(events, pool_rooms, None)
    fill_after = fill_by_room(final_events, pool_rooms, None)
    room_rows = []
    for room in sorted(pool_rooms, key=lambda r: (r.building, r.name.casefold())):
        room_rows.append(
            {
                "code": room.code, "room": room.name, "building": room.building, "kind": room.kind, "capacity": room.capacity,
                "hours_before": round(before_occ[room.code]["hours"], 1), "rate_before": round(before_occ[room.code]["rate"], 2),
                "hours_after": round(after_occ[room.code]["hours"], 1), "rate_after": round(after_occ[room.code]["rate"], 2),
                "fill_before": round(fill_before.get(room.code, {}).get("fill", 0.0), 1) if room.code in fill_before else "",
                "fill_after": round(fill_after.get(room.code, {}).get("fill", 0.0), 1) if room.code in fill_after else "",
            }
        )
    out_hours = sum(r["freed_hours"] for r in series_rows if r["status"] == "réaffecté" and r["kind"] == "amphi")
    in_hours = sum(
        r["freed_hours"] for r in series_rows if r["status"] == "réaffecté" and r["target_kind"] == "amphi"
    )
    over_rows = [
        {
            "DATE": e["date"].isoformat(), "LOCAL": e["room_name"], "BATIMENT": e["building"], "CAPACITE": e["capacity"],
            "EFFECTIF": e["effect"], "DEPASSEMENT": e["effect"] - e["capacity"], "TYPE": e["type"], "LIBELLE_MAT": e["subject"],
            "NOM_DIP": " | ".join(e["diplomas"]), "EXAMEN": "oui" if e.get("is_exam") else "non",
        }
        for e in sorted(scope, key=lambda e: (-(e["effect"] or 0) + (e["capacity"] or 0), e["date"]))
        if e["effect"] is not None and e["capacity"] and e["effect"] > e["capacity"] and not e.get("effect_estimated")
    ]
    unknown = [e for e in candidates if e["effect"] is None]
    moved_series = [r for r in series_rows if r["status"] == "réaffecté"]
    stats = {
        "scope_n": len(scope), "scope_h": sum(active_minutes(e) for e in scope) / 60,
        "candidate_n": len(candidates), "candidate_h": sum(active_minutes(e) for e in candidates) / 60,
        "exam_n": len(exam_events), "exam_h": sum(active_minutes(e) for e in exam_events) / 60,
        "other_n": len(other_events), "other_h": sum(active_minutes(e) for e in other_events) / 60,
        "unknown_n": len(unknown), "unknown_h": sum(active_minutes(e) for e in unknown) / 60,
        "series_n": len(series_rows), "moved_series": len(moved_series),
        "moved_sessions": sum(r["sessions"] for r in moved_series), "moved_hours": sum(r["hours"] for r in moved_series),
        "amphi_out_hours": out_hours, "amphi_in_hours": in_hours, "amphi_net_hours": out_hours - in_hours,
        "amphi_rate_before": sum(before_occ[r.code]["rate"] for r in amphis) / len(amphis) if amphis else 0.0,
        "amphi_rate_after": sum(after_occ[r.code]["rate"] for r in amphis) / len(amphis) if amphis else 0.0,
        "amphi_count": len(amphis),
        "fill_before": weighted_fill(events, known_ids), "fill_after": weighted_fill(final_events, known_ids),
        "conflict_pairs": conflict_pairs, "conflict_moved_sessions": len(conflict_moved),
        "conflict_moved_hours": sum(active_minutes(e) for e in events if e["id"] in conflict_moved) / 60,
        "over_n": len(over_rows), "pool_n": len(pool),
        "moved_estimated_sessions": sum(r["sessions"] for r in moved_series if r["estimated"]),
        "moved_estimated_hours": sum(r["hours"] for r in moved_series if r["estimated"]),
        "estimated_candidate_n": sum(1 for e in candidates if e.get("effect_estimated")),
        "no_room_series": sum(1 for r in series_rows if r["status"].startswith("non réaffecté")),
        "no_room_hours": sum(r["hours"] for r in series_rows if r["status"].startswith("non réaffecté")),
        "adapted_series": sum(1 for r in series_rows if r["status"].startswith("déjà adapté")),
        "over_series": sum(1 for r in series_rows if r["status"].startswith("effectif supérieur")),
    }
    return {
        "series": series_rows, "sessions": session_rows, "rooms": room_rows, "over": over_rows, "events": final_events,
        "stats": stats, "moved": moved,
    }


def hypotheses_csv_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out = []
    for row in rows:
        out.append(
            {
                "HYPOTHESE": row["key"], "DESCRIPTION": row["label"], "LOCAUX_ACCUEIL": row["rooms"],
                "SEANCES_A_REPORTER": row["total_n"], "HEURES_A_REPORTER": row["total_h"],
                "DONT_COURS_ORDINAIRES_SEANCES": row["ordinary_n"], "DONT_COURS_ORDINAIRES_HEURES": row["ordinary_h"],
                "DONT_EXAMENS_SEANCES": row["exam_n"], "DONT_EXAMENS_HEURES": row["exam_h"],
                "REPORTEES_EFFECTIF_CONNU_SEANCES": row["placed_n"], "REPORTEES_EFFECTIF_CONNU_HEURES": row["placed_h"],
                "REPORTEES_EFFECTIF_ESTIME_SEANCES": row["estimated_n"], "REPORTEES_EFFECTIF_ESTIME_HEURES": row["estimated_h"],
                "DONT_AVEC_CHEVAUCHEMENT_SEANCES": row["overlap_n"], "DONT_AVEC_CHEVAUCHEMENT_HEURES": row["overlap_h"],
                "CHEVAUCHEMENTS": row["conflicts"],
                "SANS_EFFECTIF_SEANCES": row["no_effect_n"], "SANS_EFFECTIF_HEURES": row["no_effect_h"],
                "SANS_SOLUTION_SEANCES": row["no_solution_n"], "SANS_SOLUTION_HEURES": row["no_solution_h"],
                "NON_COUVERTES_SEANCES": row["uncovered_n"], "NON_COUVERTES_HEURES": row["uncovered_h"],
                "BESOIN_NON_RESOLU_SEANCES": row["unresolved_n"], "BESOIN_NON_RESOLU_HEURES": row["unresolved_h"],
            }
        )
    return out


def classify_usage(event: dict[str, Any]) -> tuple[str, str, str]:
    """Catégorie d'usage, niveau de confiance et base du classement, d'après les intitulés et informations disponibles."""
    subject = (event["subject"] or "").casefold()
    group = " | ".join(event["diplomas"])
    course_type = (event["type"] or "").casefold()
    if group.startswith("<UTLA>") or "utla" in subject:
        return "UTLA", "élevée", "promotion <UTLA> ou intitulé UTLA"
    if re.search(r"concours|association|colloque|externe|lyc[ée]e|rectorat|forum|salon|congr[èe]s|prestataire|location", subject):
        return "Usage externe", "moyenne", "mot-clé de l'intitulé"
    if re.search(r"chaire|r[ée]union|s[ée]minaire|pr[ée]sentation|information|tuteur|info sup|scuio|atelier|conf[ée]rence", subject):
        return "Autre usage interne", "moyenne", "mot-clé de l'intitulé"
    if event.get("is_exam") or re.search(r"tiers[- ]temps", subject):
        return "Examen SSH", "moyenne", "type ou intitulé d'examen/contrôle"
    if re.search(r"r[ée]servation|r[ée]sa\b", subject) or not subject.strip():
        return "Réservation à identifier", "faible", "intitulé générique ou absent"
    if course_type in REALLOC_TYPES:
        return "Enseignement SSH", "moyenne", "type de cours (Cours/CM/TD/CTD), promotion non renseignée"
    return "Enseignement SSH", "faible", "intitulé de matière, type et promotion non renseignés"


def max_free_capacity(event: dict[str, Any], occupancy: dict[tuple[str, dt.date], list[dict[str, Any]]], recipients: list[Room]) -> int:
    """Plus grande capacité d'un local d'accueil libre au créneau de la séance (mêmes jour et horaire)."""
    best = 0
    for room in recipients:
        capacity = room.capacity or 0
        if capacity > best and all(
            not (event["start"] < other["end"] and other["start"] < event["end"]) for other in occupancy.get((room.code, event["date"]), [])
        ):
            best = capacity
    return best


def usage_free_capacities(
    unknown_events: list[dict[str, Any]], scenarios: dict[str, dict[str, Any]], recipients_by_key: dict[str, list[Room]], keys: tuple[str, ...]
) -> dict[str, dict[str, int]]:
    out: dict[str, dict[str, int]] = {event["id"]: {} for event in unknown_events}
    for key in keys:
        occupancy = room_events(scenarios[key]["events"])
        for event in unknown_events:
            out[event["id"]][key] = max_free_capacity(event, occupancy, recipients_by_key[key])
    return out


def usage_rows(unknown_events: list[dict[str, Any]], free_caps: dict[str, dict[str, int]]) -> list[dict[str, Any]]:
    rows = []
    for event in sorted(unknown_events, key=lambda e: (e["date"], e["start"], e["room_name"])):
        category, confidence, basis = classify_usage(event)
        row = {
            "ID_SEANCE": event["id"], "DATE": event["date"].isoformat(), "JOUR": WEEKDAYS_FR[event["date"].weekday()],
            "HDEBUT": f"{event['start'] // 60:02d}:{event['start'] % 60:02d}", "HFIN": f"{event['end'] // 60:02d}:{event['end'] % 60:02d}",
            "NOM_SAL": event["room_name"], "CAPACITE_LOCAL": event["capacity"], "TYPE": event["type"], "LIBELLE_MAT": event["subject"],
            "NOM_DIP": " | ".join(event["diplomas"]), "CATEGORIE": category, "CONFIANCE": confidence, "BASE_CLASSEMENT": basis,
            "HEURES_8H_18H": round(active_minutes(event) / 60, 2), "HEURES_COMPLETES": round(event["duration"] / 60, 2),
        }
        for key, cap in free_caps.get(event["id"], {}).items():
            row[f"PLUS_GRAND_LOCAL_LIBRE_{key}"] = cap
        row["EFFECTIF_A_RENSEIGNER"] = ""
        rows.append(row)
    return rows


def estimated_source_rows(
    source_events: list[dict[str, Any]], scenarios: dict[str, dict[str, Any]], keys: tuple[str, ...]
) -> list[dict[str, Any]]:
    """Séances des amphis LET placées sur effectif estimé : à vérifier en priorité."""
    rows = []
    for event in source_events:
        if not event.get("effect_estimated"):
            continue
        row = {
            "ID_SEANCE": event["id"], "DATE": event["date"].isoformat(), "HDEBUT": f"{event['start'] // 60:02d}:{event['start'] % 60:02d}",
            "HFIN": f"{event['end'] // 60:02d}:{event['end'] % 60:02d}", "AMPHI_LET": event["room_name"], "TYPE": event["type"],
            "LIBELLE_MAT": event["subject"], "NOM_DIP": " | ".join(event["diplomas"]), "EFFECTIF_ESTIME": event["effect"],
            "METHODE": event.get("effect_method", ""), "HEURES_8H_18H": round(active_minutes(event) / 60, 2),
        }
        in_conflict = False
        for key in keys:
            status = scenarios[key]["statuses"].get(event["id"], {})
            row[f"LOCAL_{key}"] = status.get("destination", "")
            room = next((r for r in scenarios[key]["placed"] if r["source_id"] == event["id"]), None) if status.get("destination") else None
            row[f"CAPACITE_{key}"] = room["capacity"] if room else ""
            row[f"STATUT_{key}"] = status.get("status", "")
            in_conflict = in_conflict or status.get("status") == "place_avec_chevauchement"
        weak = row["METHODE"].startswith("promotion parente")
        row["PRIORITE"] = "haute" if (in_conflict or weak) else ("moyenne" if row["METHODE"].startswith("groupes frères") else "faible")
        row["POURQUOI"] = ("effectif peu fiable (promotion entière) ; " if weak else "") + ("placée avec chevauchement dans une hypothèse ; " if in_conflict else "")
        rows.append(row)
    rows.sort(key=lambda r: ({"haute": 0, "moyenne": 1, "faible": 2}[r["PRIORITE"]], -r["HEURES_8H_18H"]))
    return rows


def unlocated_summary(unlocated: list[dict[str, Any]], ssh_names: set[str]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Séances SSH Pau sans salle dans l'export, susceptibles de concerner les amphis étudiés (selon l'effectif)."""
    series: dict[tuple[Any, ...], dict[str, Any]] = {}
    for item in unlocated:
        probe = {"diplomas": sorted(item["diplomas"])}
        if not item["diplomas"] or not event_is_ssh_pau(probe, ssh_names):
            continue
        effect = max(item["effects"]) if item["effects"] else None
        key = (item["subject"], item["type"], " | ".join(sorted(item["diplomas"])), item["date"].weekday(), item["start"], item["end"])
        entry = series.setdefault(
            key, {"subject": item["subject"], "type": item["type"], "groups": key[2], "weekday": WEEKDAYS_FR[key[3]], "start": item["start"],
                  "end": item["end"], "sessions": 0, "hours": 0.0, "effect": None, "exam": bool(item["is_exam"])},
        )
        entry["sessions"] += 1
        entry["hours"] += clipped_minutes(item["start"], item["end"]) / 60
        if effect is not None:
            entry["effect"] = max(entry["effect"] or 0, effect)
    rows = sorted(series.values(), key=lambda r: (-(r["effect"] or 0), -r["hours"]))
    brackets = [("Effectif inconnu", None, None), ("Moins de 50", 1, 49), ("50 à 99", 50, 99), ("100 à 149", 100, 149), ("150 et plus", 150, 10**6)]
    table = []
    for label, low, high in brackets:
        items = [r for r in rows if (r["effect"] is None if low is None else (r["effect"] is not None and low <= r["effect"] <= high))]
        table.append(
            {"bracket": label, "series": len(items), "sessions": sum(r["sessions"] for r in items), "hours": sum(r["hours"] for r in items),
             "exam_sessions": sum(r["sessions"] for r in items if r["exam"]),
             "amphi_risk": "oui" if low is not None and low >= 100 else ("à qualifier" if low is None else "non")}
        )
    return table, rows


def capacity_anomaly_rows(
    events: list[dict[str, Any]], rooms: dict[str, Room], scenarios: dict[str, dict[str, Any]], recipients_by_key: dict[str, list[Room]],
    keys: tuple[str, ...],
) -> list[dict[str, Any]]:
    """Locaux de report dont la capacité paraît douteuse (effectif observé supérieur ou très inférieur) et séances concernées."""
    codes = {r.code for key in keys for r in recipients_by_key[key]} | {r.code for r in rooms.values() if r.building == "LET" and r.kind == "amphi"}
    seen: dict[str, list[int]] = defaultdict(list)
    for event in events:
        if event["room_code"] in codes and event["effect"] is not None and not event.get("effect_estimated"):
            seen[event["room_code"]].append(event["effect"])
    rows = []
    for code in sorted(codes, key=lambda c: (rooms[c].building, rooms[c].name)):
        room = rooms[code]
        capacity = room.capacity or 0
        values = seen.get(code, [])
        over = sum(1 for v in values if v > capacity)
        top = max(values) if values else 0
        flags = []
        if capacity <= 0:
            flags.append("capacité absente ou nulle")
        if over:
            flags.append(f"{over} séances avec effectif > capacité (max {top}) : capacité peut-être sous-évaluée")
        if values and capacity >= 3 * top and capacity >= 100:
            flags.append(f"capacité au moins 3 fois l'effectif maximal observé ({top}) : à confirmer")
        if not flags:
            continue
        row = {"LOCAL": room.name, "BATIMENT": room.building, "CAPACITE": capacity, "EFFECTIF_MAX_OBSERVE": top, "ANOMALIE": " ; ".join(flags)}
        for key in keys:
            placed = [p for p in scenarios[key]["placed"] if p["room_code"] == code]
            row[f"SEANCES_LET_{key}"] = len(placed)
            row[f"DONT_JUSTES_{key}"] = sum(1 for p in placed if p["effect"] is not None and capacity and p["effect"] >= 0.8 * capacity)
        rows.append(row)
    return rows


def post_optimisation_rows(
    scenarios: dict[str, dict[str, Any]], reallocs: dict[str, dict[str, Any]], recipients_by_key: dict[str, list[Room]],
    source_events: list[dict[str, Any]], hyp_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Résultat après réaffectation : chaque séance reportée est classée une seule fois (aucun double comptage)."""
    by_hyp = {r["key"]: r for r in hyp_rows}
    sources = {e["id"]: e for e in source_events}
    output = []
    for key, result in scenarios.items():
        codes = {room.code for room in recipients_by_key[key]}
        final = reallocs[key]["events"]
        remaining = find_conflicts(final, codes)
        persistent_ids: set[str] = set()
        uncertain_ids: set[str] = set()
        persistent_min = uncertain_min = persistent_pairs = uncertain_pairs = other_pairs = 0
        suffix = f"@{key}"
        for conflict in remaining:
            minutes = max(0, min(conflict["end1"], conflict["end2"], 1080) - max(conflict["start1"], conflict["start2"], 480))
            if not (conflict["id1"].endswith(suffix) or conflict["id2"].endswith(suffix)):
                other_pairs += 1
            if conflict["unc1"] or conflict["unc2"]:
                uncertain_pairs += 1
                uncertain_min += minutes
                uncertain_ids |= {conflict["id1"], conflict["id2"]}
            else:
                persistent_pairs += 1
                persistent_min += minutes
                persistent_ids |= {conflict["id1"], conflict["id2"]}
        buckets: dict[str, list[float]] = defaultdict(lambda: [0, 0.0])
        for source_id, status in result["statuses"].items():
            event = sources[source_id]
            hours = active_minutes(event) / 60
            kind = status["status"]
            placed_id = f"{source_id}{suffix}"
            if kind == "effectif_inconnu":
                name = "not_placed_no_effect"
            elif kind in {"sans_solution", "non_couvert"}:
                name = "not_placed_other"
            elif placed_id in persistent_ids:
                name = "conflict_persistent"
            elif placed_id in uncertain_ids:
                name = "conflict_uncertain"
            else:
                name = "ok_estimated" if status.get("estimated") else "ok_known"
            buckets[name][0] += 1
            buckets[name][1] += hours
        stats = reallocs[key]["stats"]
        row: dict[str, Any] = {
            "key": key, "label": HYPOTHESIS_LABELS.get(key, key),
            "before_overlap_sessions": by_hyp[key]["overlap_n"], "before_overlap_hours": round(by_hyp[key]["overlap_h"], 1),
            "before_pairs": len(result["conflicts"]),
            "moved_for_conflict_sessions": stats["conflict_moved_sessions"], "moved_for_conflict_hours": round(stats["conflict_moved_hours"], 1),
            "persistent_pairs": persistent_pairs, "persistent_hours": round(persistent_min / 60, 1),
            "uncertain_pairs": uncertain_pairs, "uncertain_hours": round(uncertain_min / 60, 1),
            "other_pairs": other_pairs, "total": by_hyp[key]["total_n"], "total_hours": round(by_hyp[key]["total_h"], 1),
        }
        for name in ("ok_known", "ok_estimated", "conflict_persistent", "conflict_uncertain", "not_placed_no_effect", "not_placed_other"):
            row[f"{name}_n"], row[f"{name}_h"] = int(buckets[name][0]), round(buckets[name][1], 1)
        output.append(row)
    return output


def verdict_rows(post: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Conclusion par scénario : horaires inchangés après optimisation, lissage complémentaire, ou conditions à lever."""
    by_key = {r["key"]: r for r in post}
    twins = {"H1a": "H2", "H1b": "H3b", "H1c": None}
    rows = []
    for key in ("H1a", "H1b", "H1c"):
        row = by_key[key]
        twin = by_key.get(twins[key]) if twins[key] else None
        unchanged = row["conflict_persistent_n"] == 0
        smoothing = None if twin is None else (twin["persistent_pairs"] + twin["uncertain_pairs"] == 0)
        if unchanged:
            base = "Possible à horaires inchangés après optimisation des locaux (séances à effectif connu)"
        elif smoothing:
            base = f"Possible avec lissage complémentaire (simulation {twins[key]}) ; des conflits subsistent à horaires inchangés"
        elif smoothing is None:
            base = "Non démontré à horaires inchangés ; lissage de cette hypothèse non simulé"
        else:
            base = "Non démontré, même avec lissage"
        conditions = []
        if row["not_placed_no_effect_n"] or row["not_placed_other_n"]:
            conditions.append(
                f"relogement de {row['not_placed_no_effect_n'] + row['not_placed_other_n']} séances ({fr_number(row['not_placed_no_effect_h'] + row['not_placed_other_h'], 0)} h) sans effectif ou sans solution"
            )
        if row["ok_estimated_n"] or row["conflict_uncertain_n"]:
            conditions.append(f"vérification de {row['ok_estimated_n'] + row['conflict_uncertain_n']} séances placées sur effectif estimé")
        if not unchanged:
            conditions.append(f"résolution de {row['conflict_persistent_n']} séances restant en conflit")
        rows.append(
            {"key": key, "label": row["label"], "horaires_inchanges": "oui" if unchanged else "non",
             "lissage": "sans objet" if unchanged else ("non simulé" if smoothing is None else ("oui" if smoothing else "non")),
             "verdict": base, "conditions": " ; ".join(conditions) or "aucune condition identifiée"}
        )
    return rows


def verdict_csv_rows(post: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "HYPOTHESE": r["key"], "DESCRIPTION": r["label"], "SEANCES_A_REPORTER": r["total"], "HEURES_A_REPORTER": r["total_hours"],
            "CONFLITS_AVANT_SEANCES": r["before_overlap_sessions"], "CONFLITS_AVANT_HEURES": r["before_overlap_hours"],
            "RESOLUS_PAR_RELOGEMENT_SEANCES": r["moved_for_conflict_sessions"], "RESOLUS_PAR_RELOGEMENT_HEURES": r["moved_for_conflict_hours"],
            "PLACEES_SANS_CONFLIT_EFFECTIF_CONNU_SEANCES": r["ok_known_n"], "PLACEES_SANS_CONFLIT_EFFECTIF_CONNU_HEURES": r["ok_known_h"],
            "PLACEES_SANS_CONFLIT_EFFECTIF_ESTIME_SEANCES": r["ok_estimated_n"], "PLACEES_SANS_CONFLIT_EFFECTIF_ESTIME_HEURES": r["ok_estimated_h"],
            "CONFLIT_PERSISTANT_SEANCES": r["conflict_persistent_n"], "CONFLIT_PERSISTANT_HEURES": r["conflict_persistent_h"],
            "CONFLIT_INCERTAIN_SEANCES": r["conflict_uncertain_n"], "CONFLIT_INCERTAIN_HEURES": r["conflict_uncertain_h"],
            "NON_PLACEES_SANS_EFFECTIF_SEANCES": r["not_placed_no_effect_n"], "NON_PLACEES_SANS_EFFECTIF_HEURES": r["not_placed_no_effect_h"],
            "NON_PLACEES_AUTRE_SEANCES": r["not_placed_other_n"], "NON_PLACEES_AUTRE_HEURES": r["not_placed_other_h"],
            "CONFLITS_PERSISTANTS_MOMENTS": r["persistent_pairs"], "CONFLITS_PERSISTANTS_HEURES_DOUBLEES": r["persistent_hours"],
            "CONFLITS_INCERTAINS_MOMENTS": r["uncertain_pairs"], "CONFLITS_INCERTAINS_HEURES_DOUBLEES": r["uncertain_hours"],
            "CONFLITS_SANS_SEANCE_LET": r["other_pairs"],
        }
        for r in post
    ]


DEFAULT_EXAM_PERIODS = "2026-01-05:2026-01-17,2026-05-18:2026-06-30"
# Fermetures de l'université 2025-2026 (jours ouvrés sans cours) ; les jours fériés nationaux sont déjà retranchés par ailleurs.
# Une vacance « du vendredi soir au lundi matin » ne retire que les jours ouvrés strictement entre ces deux dates.
DEFAULT_CLOSURE_PERIODS = (
    "2025-10-27:2025-10-31,"  # Toussaint
    "2025-12-22:2026-01-02,"  # Noël
    "2026-02-16:2026-02-20,"  # Hiver
    "2026-04-06:2026-04-17"  # Printemps
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="ETL d'occupation des salles UPPA : produit uniquement les CSV (aucun HTML).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--dossier", type=Path, default=Path(__file__).resolve().parent / "BDD", help="Dossier des sources.")
    parser.add_argument("--planning", help="Chemin vers le CSV de planning.")
    parser.add_argument("--salles", help="Chemin vers EXP_SALLE (catalogue des salles).")
    parser.add_argument("--promotions", help="Chemin vers EXP_PROMOTION.")
    parser.add_argument("--surfaces", type=Path, default=None, help="Classeur SURFACES GENERAL (Ville/Site/Code Bât) ; BDD/SURFACES*.xlsx par défaut.")
    parser.add_argument("--fermetures", default=DEFAULT_CLOSURE_PERIODS, help="Périodes de fermeture AAAA-MM-JJ:AAAA-MM-JJ séparées par des virgules ; chaîne vide pour aucune.")
    parser.add_argument("--examens", default=DEFAULT_EXAM_PERIODS, help="Périodes d'examen AAAA-MM-JJ:AAAA-MM-JJ séparées par des virgules ; chaîne vide pour aucune.")
    parser.add_argument("--sortie", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--debut", help="Date ISO YYYY-MM-DD incluse.")
    parser.add_argument("--fin", help="Date ISO YYYY-MM-DD incluse.")
    parser.add_argument("--exclure-semaines", help="Semaines ISO à exclure, ex. 43,52,1.")
    parser.add_argument("--modele-effectifs", action="store_true", help="Écrit les NOM_DIP distincts à renseigner puis s'arrête.")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
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
        excluded_weeks = {int(v.strip()) for v in args.exclure_semaines.split(",") if v.strip()} if args.exclure_semaines else set()
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
        # --- Extraction : catalogue des salles, promotions, planning ---
        rooms, room_notes = load_rooms(room_path)
        promotion_iterator = iter_csv_dicts(promotion_path)
        promotion_metadata = next(promotion_iterator)
        if not PROMOTION_REQUIRED.issubset(set(promotion_metadata["__headers__"])):
            raise SourceError(f"Colonnes requises absentes de {promotion_path.name}.")
        promotion_effects: dict[str, set[int]] = defaultdict(set)
        promotion_names: set[str] = set()
        ssh_pau_names: set[str] = set()
        for row in promotion_iterator:
            name = row.get("NOM_DIP", "").strip()
            if name:
                promotion_names.add(name)
                if "<Structure>SSH - Pau" in str(row.get("FAMILLE", "")):
                    ssh_pau_names.add(name)
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
        active_dates = [d for d in schedule["active_dates"] if start <= d <= end]
        schedule["events"] = [e for e in schedule["events"] if start <= e["date"] <= end]
        if not active_dates:
            raise AnalysisError("Aucun jour pédagogique ouvré ne reste après les filtres.")
        events = schedule["events"]
        source_codes = {r.code for r in rooms.values() if r.building == "LET" and r.kind == "amphi"}
        schedule["missing_effect"] = sum(e["effect"] is None for e in events)
        estimate_methods: Counter[str] = Counter()  # pas d'estimation d'effectifs dans cet ETL
        source_events = [e for e in events if e["room_code"] in source_codes]
        schedule["estimated_sources"] = sum(1 for e in source_events if e.get("effect_estimated"))
        schedule["estimated_all"] = sum(estimate_methods.values())
        schedule["estimate_methods"] = dict(estimate_methods)
        schedule["missing_effect_source"] = sum(e["effect"] is None for e in source_events)
        print(f"Séances sans effectif retrouvé : {schedule['missing_effect']} localisées ; {schedule['missing_effect_source']} dans les amphis LET à reporter.")

        if args.modele_effectifs:
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            destination = directory / f"effectifs_a_renseigner_{stamp}.csv"
            count = model_template_from_stream(schedule_path, destination)
            print(f"Modèle créé : {destination} ({count} NOM_DIP distincts).")
            return 0

        surfaces_path = args.surfaces.expanduser().resolve() if args.surfaces else next(iter(sorted(directory.glob("SURFACES*.xlsx"))), None)
        building_sites = load_building_sites(surfaces_path) if surfaces_path and surfaces_path.is_file() else {}

        # --- Transformation : locaux d'accueil par hypothèse ---
        recipients_h1 = [r for r in rooms.values() if r.building == "DEG" and r.kind == "amphi"]
        recipients_pau = recipients_h1 + [
            r for r in rooms.values()
            if r.kind == "amphi" and r.building not in {"DEG", "LET"} and (r.capacity or 0) > 0
            and building_sites.get(r.building, {}).get("site") == "CAMPUS DE PAU"
        ]
        recipients_h3 = [
            r for r in rooms.values()
            if (r.building == "DEG" and r.kind == "amphi") or (r.building in {"LET", "DEG"} and r.kind == "grande_salle")
        ]
        if len(recipients_h1) != 5:
            raise AnalysisError(f"5 amphithéâtres DEG attendus dans EXP_SALLE, {len(recipients_h1)} trouvés.")
        if len([r for r in rooms.values() if r.building == "LET" and r.kind == "amphi"]) != 3:
            raise AnalysisError("Le catalogue ne contient pas exactement trois amphithéâtres LET identifiables.")
        if not recipients_h3:
            raise AnalysisError("Aucun local d'accueil H3 identifié dans LET/DEG.")

        source_ids = {e["id"] for e in source_events}
        baseline_without_source = [e for e in events if e["id"] not in source_ids]
        recipients_by_key = {"H1a": recipients_h1, "H1b": recipients_h3, "H1c": recipients_pau, "H2": recipients_h1, "H3b": recipients_h3}
        flexible_keys = {"H2", "H3b"}
        scenarios: dict[str, dict] = {}
        for key, recipients in recipients_by_key.items():
            scenarios[key] = scenario_result(key, baseline_without_source, source_events, recipients, active_dates, key in flexible_keys)
            validate_scenario(key, scenarios[key], source_events, recipients, key in flexible_keys)

        # --- Chargement : uniquement des CSV ---
        run_directory, stamp, created_at = create_run_directory(args.sortie.expanduser().resolve())
        written: list[str] = []

        def out(name: str, headers: list[str], rows: list[dict]) -> None:
            write_csv(run_directory / f"{name}_{stamp}.csv", headers, rows)
            written.append(f"{name}_{stamp}.csv")

        over_rows = overdimension_rows(events, rooms)
        out("06_surdimensionnement", ["date", "room", "code", "building", "capacity", "effect", "unused", "type", "subject", "diplomas", "seriesOccurrences", "replacement", "replacementCode"], over_rows)
        details = detail_rows(events, scenarios, rooms)
        out("07_detail_seances", list(details[0]) if details else ["ID_SEANCE", "DATE", "CODE_SAL"], details)

        indicators = build_indicators(events, rooms, scenarios, active_dates, exam_periods)
        anomalies = anomaly_rows(schedule, events)
        out("14_anomalies", ["type", "valeur", "nombre", "explication"], anomalies)
        indicator_files = indicator_csvs(run_directory, stamp, indicators, rooms)
        written += [p.name for p in indicator_files]

        hyp_rows = hypothesis_rows(scenarios, source_events, recipients_by_key)
        hyp_csv = hypotheses_csv_rows(hyp_rows)
        out("18_hypotheses_report", list(hyp_csv[0]), hyp_csv)

        pool = reallocation_pool(rooms, building_sites)
        reallocs = {"Référence": simulate_global_reallocation(events, rooms, pool, ssh_pau_names, active_dates)}
        for key, result in scenarios.items():
            reallocs[key] = simulate_global_reallocation(result["events"], rooms, pool, ssh_pau_names, active_dates)
        realloc = reallocs["Référence"]
        series_headers = ["estimated", "room", "code", "building", "kind", "capacity", "max_effect", "mean_effect", "weekday", "start", "end", "subject", "type", "diplomas", "sessions", "hours", "target", "target_code", "target_capacity", "target_building", "target_kind", "freed_hours", "status"]
        session_headers = list(realloc["sessions"][0]) if realloc["sessions"] else ["ID_SEANCE"]
        room_headers = ["code", "room", "building", "kind", "capacity", "hours_before", "rate_before", "hours_after", "rate_after", "fill_before", "fill_after"]
        over_headers = ["DATE", "LOCAL", "BATIMENT", "CAPACITE", "EFFECTIF", "DEPASSEMENT", "TYPE", "LIBELLE_MAT", "NOM_DIP", "EXAMEN"]
        for key, item in reallocs.items():
            suffix = "reference" if key == "Référence" else key
            out(f"19_reaffectation_avant_apres_{suffix}", series_headers, item["series"])
            out(f"20_reaffectation_seances_{suffix}", session_headers, item["sessions"])
            out(f"21_reaffectation_locaux_{suffix}", room_headers, item["rooms"])
            out(f"22_effectif_superieur_capacite_{suffix}", over_headers, item["over"])

        keys3 = ("H1a", "H1b", "H1c")
        unknown_events = [e for e in source_events if e["effect"] is None]
        free_caps = usage_free_capacities(unknown_events, scenarios, recipients_by_key, keys3)
        u_rows = usage_rows(unknown_events, free_caps)
        out("17_usages_sans_effectif", list(u_rows[0]) if u_rows else ["ID_SEANCE"], u_rows)

        est_rows = estimated_source_rows(source_events, scenarios, keys3)
        if est_rows:
            out("23_verification_effectifs_estimes_LET", list(est_rows[0]), est_rows)
        _, unlocated_rows = unlocated_summary(schedule["unlocated"], ssh_pau_names)
        out("24_seances_sans_salle_SSH_Pau", ["subject", "type", "groups", "weekday", "start", "end", "sessions", "hours", "effect", "exam"], unlocated_rows)
        anomalies_cap = capacity_anomaly_rows(events, rooms, scenarios, recipients_by_key, keys3)
        if anomalies_cap:
            out("25_anomalies_capacite", list(anomalies_cap[0]), anomalies_cap)
        post = post_optimisation_rows(scenarios, reallocs, recipients_by_key, source_events, hyp_rows)
        post_csv = verdict_csv_rows(post)
        out("26_apres_optimisation", list(post_csv[0]), post_csv)
        verdicts = verdict_rows(post)
        out("27_conclusion_scenarios", list(verdicts[0]), verdicts)

        print(f"Dossier de résultats : {run_directory}")
        print(f"{len(written)} CSV écrits (aucun HTML) :")
        for name in sorted(written):
            print(f"  {name}")
        print(f"Jours pédagogiques retenus : {len(active_dates)}.")
        return 0
    except (OSError, csv.Error, SourceError, AnalysisError, ValueError) as error:
        print(f"ARRÊT : {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
