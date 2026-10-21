#!/usr/bin/env python3
"""Sépare le dashboard consolidé 08_dashboard_complet_<horodatage>.html en fichiers distincts.

Le dashboard d'origine est un seul fichier (≈ 27 Mo) : la page d'accueil et les 13 vues (iframe) y embarquent leur HTML,
leur CSS, leur JavaScript et leurs données. Ce script écrit, à côté de lui, un dossier dashboard_<horodatage>/ :

    index.html            page d'accueil (structure HTML uniquement)
    pages/<vue>.html      une page HTML par vue (structure uniquement)
    css/*.css             feuilles de style (accueil, vues, correctif de contraste)
    js/*.js               logique JavaScript (accueil, vues, info-bulles)
    data/*.js             données (rawData, heatmaps, hebdomadaire, conflits, textes des cas, payloads)

Les données sont des fichiers .js (const nom = {...};) et non des .json : un navigateur n'autorise pas fetch() sur un
fichier local (file://), alors qu'une balise <script src> fonctionne en double-clic. Ouvrir dashboard_<horodatage>/index.html.

Usage :
    python separer_dashboard.py resultats_AAAAMMJJ_HHMMSS_xxxxxx/08_dashboard_complet_AAAAMMJJ_HHMMSS_xxxxxx.html
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

DECODER = json.JSONDecoder()
BLOCK = re.compile(r"<(style|script)([^>]*)>(.*?)</\1>", re.S)
MIN_DATA_CHARS = 2000  # une constante JSON plus courte reste dans le JavaScript (réglages, libellés)


class Assets:
    """Écrit les fichiers extraits ; un contenu identique n'est écrit qu'une fois (info-bulles, correctif de contraste)."""

    def __init__(self, root: Path) -> None:
        self.root = root
        self.by_hash: dict[str, str] = {}
        self.used: set[str] = set()

    def write(self, folder: str, name: str, ext: str, content: str) -> str:
        digest = hashlib.sha1((folder + content).encode("utf-8")).hexdigest()
        if digest in self.by_hash:
            return self.by_hash[digest]
        rel, n = f"{folder}/{name}.{ext}", 2
        while rel in self.used:
            rel, n = f"{folder}/{name}_{n}.{ext}", n + 1
        self.used.add(rel)
        path = self.root / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
        self.by_hash[digest] = rel
        return rel


def pull_data_constants(script: str) -> tuple[str, str]:
    """Retire du script les constantes `const nom = <JSON volumineux>;` et les renvoie à part (données)."""
    data, out, pos = [], [], 0
    for match in re.finditer(r"\b(?:const|let|var)\s+(\w+)\s*=\s*", script):
        if match.start() < pos:
            continue
        try:
            _, end = DECODER.raw_decode(script, match.end())
        except ValueError:
            continue
        if end - match.end() < MIN_DATA_CHARS:
            continue
        tail = end
        while tail < len(script) and script[tail] in " \t\r\n":
            tail += 1
        if tail >= len(script) or script[tail] != ";":
            continue  # pas une déclaration de premier niveau terminée par « ; »
        out.append(script[pos:match.start()])
        data.append(f"const {match.group(1)} = {script[match.end():end]};")
        pos = tail + 1
    out.append(script[pos:])
    return "".join(out), "\n".join(data)


def split_page(html: str, key: str, assets: Assets, depth: int) -> str:
    """Remplace styles, scripts et données d'une page par des liens vers des fichiers ; renvoie le HTML restant."""
    up = "../" * depth
    counters = {"style": 0, "script": 0}

    def replace(match: re.Match) -> str:
        tag, attrs, body = match.group(1), match.group(2), match.group(3)
        ident = re.search(r'id="([\w-]+)"', attrs)
        name = (ident.group(1) if ident else key).replace("-", "_")
        if tag == "style":
            counters["style"] += 1
            rel = assets.write("css", name if ident else f"{key}_{counters['style']}", "css", body)
            return f'<link rel="stylesheet" href="{up}{rel}">'
        if "application/json" in attrs:  # données JSON lues par JSON.parse(...getElementById(...).textContent)
            rel = assets.write("data", f"{key}_{name}", "js", f"window.DATA_{name.upper()} = {body};")
            return f'<script src="{up}{rel}"></script>'
        counters["script"] += 1
        logic, data = pull_data_constants(body)
        for ident_json in re.findall(r"JSON\.parse\(document\.getElementById\('(\w+)'\)\.textContent\)", logic):
            logic = logic.replace(f"JSON.parse(document.getElementById('{ident_json}').textContent)", f"window.DATA_{ident_json.upper()}")
        tags = ""
        if data:
            tags += f'<script src="{up}{assets.write("data", key, "js", data)}"></script>'
        rel = assets.write("js", name if ident else (key if counters["script"] == 1 else f"{key}_{counters['script']}"), "js", logic)
        return tags + f'<script src="{up}{rel}"></script>'

    return BLOCK.sub(replace, html)


def split_dashboard(html_text: str, out_dir: Path) -> Path:
    """Écrit le dashboard séparé dans out_dir (index.html, pages/, css/, js/, data/) et renvoie index.html."""
    pages_match = re.search(r'<script type="application/json" id="embedded-pages">(.*?)</script>', html_text, re.S)
    cases_match = re.search(r'<script type="application/json" id="case-texts">(.*?)</script>', html_text, re.S)
    if not pages_match or not cases_match:
        raise ValueError("Structure inattendue : blocs embedded-pages / case-texts introuvables.")
    pages = json.loads(pages_match.group(1).replace("<\\/", "</"))
    cases = cases_match.group(1).replace("<\\/", "</")
    assets = Assets(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Page d'accueil : les deux blocs JSON sortent en données, l'iframe charge pages/<vue>.html au lieu de srcdoc.
    shell = html_text.replace(pages_match.group(0), "").replace(cases_match.group(0), "<!--CASES-->").replace(' srcdoc=""', "")
    shell = re.sub(r'<a href="([^"#:/]+\.csv)"', lambda m: f'<a href="../{m.group(1)}"', shell)  # les CSV restent dans le dossier parent
    cases_rel = assets.write("data", "cases", "js", f"window.DATA_CASE_TEXTS = {cases};")
    shell = shell.replace("<!--CASES-->", f'<script src="{cases_rel}"></script>')

    def home_script(match: re.Match) -> str:
        if "application/json" in match.group(2) or " src=" in match.group(2):
            return match.group(0)
        body = match.group(3)
        body = body.replace("pages=JSON.parse(document.getElementById('embedded-pages').textContent),", "")
        body = body.replace("cases=JSON.parse(document.getElementById('case-texts').textContent)", "cases=window.DATA_CASE_TEXTS")
        body = body.replace("frame.srcdoc=pages[key]", "frame.src='pages/'+key+'.html'")
        return f'<script src="{assets.write("js", "dashboard", "js", body)}"></script>'

    shell = re.sub(r"<(script)([^>]*)>(.*?)</\1>", home_script, shell, flags=re.S)
    shell = re.sub(r"<(style)([^>]*)>(.*?)</\1>", lambda m: f'<link rel="stylesheet" href="{assets.write("css", "dashboard", "css", m.group(3))}">', shell, flags=re.S)
    (out_dir / "index.html").write_text(shell, encoding="utf-8")

    for key, page in pages.items():
        target = out_dir / "pages" / f"{key}.html"
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(split_page(page, key, assets, depth=1), encoding="utf-8")
    return out_dir / "index.html"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("dashboard", type=Path, help="08_dashboard_complet_<horodatage>.html")
    parser.add_argument("--sortie", type=Path, help="Dossier de sortie (par défaut dashboard_<horodatage>/ à côté du fichier)")
    args = parser.parse_args(argv)
    source = args.dashboard.expanduser().resolve()
    stamp = source.stem.replace("08_dashboard_complet_", "")
    out = (args.sortie or source.parent / f"dashboard_{stamp}").expanduser().resolve()
    try:
        index = split_dashboard(source.read_text(encoding="utf-8"), out)
    except (OSError, ValueError) as error:
        print(f"ARRÊT : {error}", file=sys.stderr)
        return 2
    total = sum(p.stat().st_size for p in out.rglob("*") if p.is_file())
    print(f"Dashboard séparé : {index} ({sum(1 for p in out.rglob('*') if p.is_file())} fichiers, {total / 1e6:.1f} Mo)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
