import io
import os
import re
import tempfile
import unicodedata
import statistics
from dataclasses import dataclass
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path

import fitz
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

PARSER_VERSION = "4.0-regression-tested"


# ============================================================
# MODELLI DATI
# ============================================================

@dataclass
class TeamInfo:
    name: str
    locality: str = ""
    address: str = ""
    time: str = ""
    field_name: str = ""
    day: str = ""


@dataclass
class Match:
    date: str
    home: str
    away: str
    time: str = ""
    locality: str = ""
    address: str = ""
    round_no: str = ""


@dataclass
class Section:
    competition: str
    group: str
    teams: dict
    matches: list
    source_format: str

    @property
    def label(self):
        bits = [x for x in [self.competition, (f"GIRONE {self.group}" if self.group else "")] if x]
        return " - ".join(bits) if bits else "CALENDARIO"


# ============================================================
# NORMALIZZAZIONE BASE
# ============================================================

def clean(value):
    value = (value or "").replace("\u00a0", " ")
    value = unicodedata.normalize("NFKD", value)
    value = "".join(c for c in value if not unicodedata.combining(c))
    value = value.upper().replace("’", "'").replace("`", "'")
    value = re.sub(r"\s+", " ", value).strip(" |\t\r\n")
    return value


def normalize_time(value):
    s = clean(value).replace(".", ":")
    m = re.search(r"\b([0-2]?\d):([0-5]\d)\b", s)
    if not m:
        return ""
    h = int(m.group(1))
    if h > 23:
        return ""
    return f"{h:02d}:{m.group(2)}"


def normalize_numeric_date(value):
    m = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b", value or "")
    if not m:
        return ""
    d, mo, y = map(int, m.groups())
    if y < 100:
        y += 2000
    try:
        return datetime(y, mo, d).strftime("%d/%m/%Y")
    except ValueError:
        return ""


def sort_date_value(value):
    try:
        return datetime.strptime(value, "%d/%m/%Y")
    except Exception:
        return datetime.max


# ============================================================
# LOCALITA' / INDIRIZZI
# ============================================================

LEGAL_PATTERNS = [
    r"S\.?S\.?D\.?\s*A\.?\s*R\.?\s*L\.?$", r"SSD\s*A\s*R\s*L$", r"S\.?S\.?D\.?$",
    r"SSD(?:ARL|RL)?$", r"A\.?S\.?D\.?$", r"ASD$", r"U\.?S\.?D\.?$", r"USD$",
    r"G\.?S\.?D\.?$", r"GSD$", r"A\.?D\.?P\.?$", r"ADP$", r"S\.?R\.?L\.?$", r"SRL$",
    r"A\.?\s*R\.?\s*L\.?$", r"ARL$", r"C\.?V\.?$", r"F\.?B\.?C\.?$", r"POL\.?D\.?$", r"POL\.?$",
]


def split_squad_suffix(name):
    """Riconosce SQ.B anche se attaccato: BREGNANESESQ.B / A.S.DSQ.B."""
    s = clean(name)
    m = re.search(r"SQ\.?\s*([A-Z])\b", s)
    suffix = f"SQ{m.group(1)}" if m else ""
    s = re.sub(r"SQ\.?\s*[A-Z]\b", "", s).strip()
    return s, suffix


def strip_legal_suffix(name):
    s, sq = split_squad_suffix(name)
    # A.C.D. davanti al nome è una forma societaria, non parte del nome.
    s = re.sub(r"^(?:A\.?\s*C\.?\s*D\.?|A\.?C\.?D\.?)\s*", "", s).strip()
    changed = True
    while changed:
        changed = False
        for pattern in LEGAL_PATTERNS:
            ns = re.sub(r"\s*\b" + pattern, "", s).strip(" .-")
            if ns != s:
                s = ns
                changed = True
    return s, sq


def extract_locality_from_field(field, team_name=""):
    """Estrae la località dalla colonna Campo/Località, inclusi PDF senza ' - '."""
    s = clean(field)
    if not s:
        return ""

    # Separatore esplicito con o senza spazi.
    parts = [p.strip() for p in re.split(r"\s*-\s*", s) if p.strip()]
    if len(parts) >= 2:
        tail = parts[-1]
        tail = re.sub(r"^CAMPO\s+(?:N\.?\s*)?[A-Z0-9]+\s+", "", tail).strip()
        return tail

    # Località dopo (E.A), (E.A.) ecc.
    m = re.search(r"\)\s*([A-Z0-9' .]+)$", s)
    if m:
        tail = clean(m.group(1)).strip(" -")
        if tail:
            return tail

    # Località dopo E.A. non racchiuso tra parentesi.
    m = re.search(r"\bE\.?\s*A\.?\s*([A-Z0-9' .]+)$", s)
    if m:
        tail = clean(m.group(1)).strip(" -")
        if tail:
            return tail

    # Campo 1 / Campo N.2 / Campo A + località.
    m = re.search(r"\bCAMPO\s+(?:N\.?\s*)?[A-Z0-9]+\s+(.+)$", s)
    if m:
        return clean(m.group(1)).strip(" -")

    # N.2 + località.
    m = re.search(r"\bN\.?\s*\d+\s+(.+)$", s)
    if m:
        return clean(m.group(1)).strip(" -")

    # Forme semplici: C.S.COMUNALE GERENZANO / C.S.COMUNALE ROVELLASCA.
    if '"' not in s and "(" not in s:
        m = re.match(r"^(?:C\.?\s*S\.?\s*)?(?:CENTRO SPORTIVO\s+)?(?:COMUNALE|PARROCCHIALE)\s+(.+)$", s)
        if m:
            return clean(m.group(1)).strip(" -")

    # Ultimo fallback: se il campo termina esattamente col nome squadra pulito.
    base, _ = strip_legal_suffix(team_name)
    base = clean(base)
    if base and s.endswith(base):
        return base

    return ""


def normalize_locality_for_excel(value):
    s = clean(value)
    if not s:
        return ""
    s = re.sub(r"\bLOC\.\s*", "LOC. ", s)
    s = re.sub(r"\bFRAZ\.\s*", "FRAZ. ", s)
    s = re.sub(r"\s*,\s*", ", ", s)
    s = re.sub(r"\s+", " ", s).strip(" ,-")
    return s


def normalize_street_address(value, locality=""):
    s = clean(value)
    loc = normalize_locality_for_excel(locality)
    if not s:
        return ""

    replacements = [
        (r"^\s*P\.?\s*ZA\.?\s+", "PIAZZA "),
        (r"^\s*P\.?\s*ZZA\.?\s+", "PIAZZA "),
        (r"^\s*P\.?\s*LE\.?\s+", "PIAZZALE "),
        (r"^\s*V\.?\s*LE\.?\s+", "VIALE "),
        (r"^\s*C\.?\s*SO\.?\s+", "CORSO "),
        (r"^\s*L\.?\s*GO\.?\s+", "LARGO "),
    ]
    for pattern, replacement in replacements:
        s = re.sub(pattern, replacement, s)

    s = re.sub(r"\bANG\.\s*", "ANG. ", s)
    s = re.sub(r"\bLOC\.\s*", "LOC. ", s)
    s = re.sub(r"\bFRAZ\.\s*", "FRAZ. ", s)

    # ROMA11 -> ROMA 11
    s = re.sub(r"([A-Z])(?=\d{1,4}(?:/[A-Z0-9]+)?\b)", r"\1 ", s)

    # Rimuove la località se accidentalmente già in fondo all'indirizzo.
    if loc:
        s = re.sub(rf"\s*,?\s*{re.escape(loc)}\s*$", "", s, flags=re.I).strip()

    # N.11 / N°11 / N 11 / NUM.11 -> 11
    s = re.sub(
        r"\s+(?:N\.?|N°|NR\.?|NUM\.?)\s*(\d+[A-Z]?(?:[/\-]\d+[A-Z]?)?(?:/[A-Z])?)\b",
        r" \1",
        s,
    )

    s = re.sub(r"\s*,\s*", ", ", s)
    s = re.sub(r"\s+", " ", s).strip(" ,")
    s = re.sub(r"\bS\.?\s*N\.?\s*C\.?\b", "SNC", s)

    # Civico finale: 11, 38/A, 9/11, 23/25, 14 A, 10/B.
    m = re.match(
        r"^(.*?)(?:,\s*|\s+)(\d+(?:[/\-]\d+)?(?:/[A-Z])?|\d+\s+[A-Z])"
        r"(\s*(?:\([^)]*\)|\"[^\"]*\")\s*)?$",
        s,
    )
    if m:
        street = m.group(1).strip(" ,")
        civic = re.sub(r"\s+", "", m.group(2).strip())
        suffix = (m.group(3) or "").strip()
        if street:
            s = f"{street}, {civic}"
            if suffix:
                s += f" {suffix}"

    s = re.sub(r",\s*(\d+(?:[/\-]\d+)?(?:/[A-Z])?)\b", r", \1", s)
    s = re.sub(r",\s*SNC\b", " SNC", s)
    return s.strip(" ,-")


def indirizzo_excel(match):
    loc = normalize_locality_for_excel(match.locality)
    addr = normalize_street_address(match.address, loc)
    if addr and loc:
        return f"{addr} - {loc}"
    return addr or loc


# ============================================================
# TABELLA SOCIETA' / CAMPI
# ============================================================

def parse_modern_field_table_grid(page):
    """Legge la griglia reale PDF anche con pagina ruotata e colonne sparse."""
    try:
        tables = page.find_tables().tables
    except Exception:
        return {}

    best = {}

    for table in tables:
        try:
            data = table.extract()
        except Exception:
            continue
        if not data or len(data) < 2:
            continue

        header_idx = None
        for ri, row in enumerate(data[:4]):
            vals = [clean(c or "") for c in row]
            joined = " | ".join(vals)
            if "SOCIETA" in joined and "CAMPO" in joined and "INDIRIZZO" in joined:
                header_idx = ri
                break
        if header_idx is None:
            continue

        teams = {}
        for row in data[header_idx + 1:]:
            vals = [clean(c or "") if c is not None else "" for c in row]
            nonempty = [(i, v) for i, v in enumerate(vals) if v]
            if not nonempty:
                continue

            # Individua il codice campo numerico: primo intero 1..5 cifre dopo il nome.
            code_pos = next((i for i, v in nonempty if re.fullmatch(r"\d{1,5}", v)), None)
            if code_pos is None:
                continue

            name = clean(" ".join(v for i, v in nonempty if i < code_pos))
            if not name or name.startswith("SOCIETA") or name.startswith("LA SOCIETA"):
                continue

            time_pos = None
            for i, v in nonempty:
                if normalize_time(v):
                    time_pos = i
            day_pos = next((i for i, v in nonempty if v in {"SABATO", "DOMENICA", "VENERDI", "VENERDÌ"}), None)

            middle = [
                (i, v) for i, v in nonempty
                if i > code_pos and (time_pos is None or i < time_pos) and (day_pos is None or i < day_pos)
            ]
            if not middle:
                continue

            fieldloc = middle[0][1]
            address = middle[1][1] if len(middle) > 1 else ""
            tm = normalize_time(vals[time_pos]) if time_pos is not None else ""
            day = vals[day_pos].title() if day_pos is not None else ""

            teams[name] = TeamInfo(
                name=name,
                locality=extract_locality_from_field(fieldloc, name),
                address=address,
                time=tm,
                field_name=fieldloc,
                day=day,
            )

        if len(teams) > len(best):
            best = teams

    return best


def _group_word_lines(words, tol=2.5):
    rows = []
    for w in sorted(words, key=lambda x: ((x[1] + x[3]) / 2, x[0])):
        yc = (w[1] + w[3]) / 2
        if not rows or abs(yc - rows[-1][0]) > tol:
            rows.append([yc, [w]])
        else:
            rows[-1][1].append(w)
    return rows


def parse_modern_field_table_words_fallback(page):
    """Fallback a coordinate se find_tables non è disponibile."""
    words = page.get_text("words") or []
    soc = [w for w in words if clean(w[4]).startswith("SOCIETA")]
    if not soc:
        return {}
    hdr_y = min(w[1] for w in soc)
    code_headers = [w for w in words if clean(w[4]) in {"N.", "N"} and abs(w[1] - hdr_y) < 5]
    if not code_headers:
        return {}
    cx = code_headers[0][0]

    teams = {}
    data_words = [w for w in words if w[1] > hdr_y + 5 and w[1] < page.rect.height * .92]
    for _, ws in _group_word_lines(data_words):
        ordered = sorted(ws, key=lambda w: w[0])
        codes = [w for w in ordered if re.fullmatch(r"\d{1,5}", clean(w[4])) and abs(w[0] - cx) < 80]
        if not codes:
            continue
        code = min(codes, key=lambda w: abs(w[0] - cx))
        name = clean(" ".join(w[4] for w in ordered if w[2] <= code[0] - 1))
        if not name:
            continue
        # Questo fallback è deliberatamente conservativo: se manca la griglia,
        # almeno consente il riconoscimento del nome squadra.
        teams[name] = TeamInfo(name=name)
    return teams


def parse_team_table(page):
    teams = parse_modern_field_table_grid(page)
    if len(teams) >= 4:
        return teams
    fallback = parse_modern_field_table_words_fallback(page)
    return fallback if len(fallback) > len(teams) else teams


# ============================================================
# MATCH NOMI SQUADRA / ABBREVIAZIONI
# ============================================================

def _team_tokens(name):
    s, sq = strip_legal_suffix(name)
    tokens = re.findall(r"[A-Z0-9]+", clean(s))
    tokens = [re.sub(r"^(\d{3,4})[A-Z]$", r"\1", x) for x in tokens]
    return tokens, sq


def _token_compat(query_token, candidate_token):
    if query_token == candidate_token:
        return 1.0
    if query_token in {"AC", "ACC"} and candidate_token in {"ACADEMY", "ACCADEMIA"}:
        return .99
    if query_token == "S" and candidate_token == "SAN":
        return .97
    if len(query_token) == 1 and candidate_token.startswith(query_token):
        return .90
    if len(query_token) >= 2 and candidate_token.startswith(query_token):
        return .96
    return 0.0


def _sequence_abbrev_score(query, candidate):
    qt, qs = _team_tokens(query)
    ct, cs = _team_tokens(candidate)

    if qs:
        if qs != cs:
            return 0.0
    elif cs:
        # Evita COMO 1907 -> COMO 1907 SQ.B quando entrambe esistono.
        return 0.0

    if not qt or not ct:
        return 0.0

    m, n = len(qt), len(ct)
    dp = [[-10**9] * (n + 1) for _ in range(m + 1)]
    dp[0][0] = 0.0

    for i in range(m + 1):
        for j in range(n + 1):
            current = dp[i][j]
            if current < -10**8:
                continue
            if j < n:
                penalty = .15 if ct[j] in {"GS", "FC", "SC", "US", "POL", "SPORT", "SPORTIVA"} else .35
                dp[i][j + 1] = max(dp[i][j + 1], current - penalty)
            if i < m and j < n:
                comp = _token_compat(qt[i], ct[j])
                if comp:
                    dp[i + 1][j + 1] = max(dp[i + 1][j + 1], current + comp)

    return max(dp[m]) / max(1, m)


def _leading_acronym(name):
    c = clean(name)
    prefix = re.match(r"^\s*((?:[A-Z]\.\s*){2,8})", c)
    if not prefix:
        return ""
    return "".join(re.findall(r"([A-Z])\.", prefix.group(1)))


def _candidate_initials(name):
    tokens, _ = _team_tokens(name)
    return "".join(x[0] for x in tokens if x)


def smart_canonical_team(name, teams):
    c = clean(name)
    if "RIPOSA" in c or "RIPOSO" in c:
        return "RIPOSA"

    q_base, q_sq = strip_legal_suffix(name)
    q_key = re.sub(r"[^A-Z0-9]+", "", clean(q_base))

    exact = []
    for team in teams:
        t_base, t_sq = strip_legal_suffix(team)
        t_key = re.sub(r"[^A-Z0-9]+", "", clean(t_base))
        if q_key == t_key and q_sq == t_sq:
            exact.append(team)
    if len(exact) == 1:
        return exact[0]

    # F.O.R. -> FALOPPIESE OLGIATE RONAGO; preserva SQ.B se presente.
    acronym = _leading_acronym(name)
    if acronym:
        candidates = []
        for team in teams:
            _, t_sq = strip_legal_suffix(team)
            if q_sq and t_sq != q_sq:
                continue
            if not q_sq and t_sq:
                continue
            if _candidate_initials(team).startswith(acronym):
                candidates.append(team)
        if len(candidates) == 1:
            return candidates[0]

    scored = []
    for team in teams:
        score = _sequence_abbrev_score(name, team)
        t_base, t_sq = strip_legal_suffix(team)
        if (q_sq and q_sq == t_sq) or (not q_sq and not t_sq):
            t_key = re.sub(r"[^A-Z0-9]+", "", clean(t_base))
            similarity = SequenceMatcher(None, q_key, t_key).ratio()
            if q_key and t_key and (q_key in t_key or t_key in q_key):
                similarity = max(similarity, .88)
            score = max(score, similarity)
        scored.append((score, team))

    scored.sort(reverse=True)
    if scored and scored[0][0] >= .62:
        if len(scored) == 1 or scored[0][0] - scored[1][0] >= .04 or scored[0][0] >= .90:
            return scored[0][1]
    return None


# ============================================================
# CALENDARIO GRAFICO MODERNO
# ============================================================

def detect_modern_round_headers(page):
    words = page.get_text("words") or []
    headers = []
    for w in words:
        if clean(w[4]) != "GIORNATA":
            continue
        yc = (w[1] + w[3]) / 2
        candidates = [
            q for q in words
            if re.fullmatch(r"\d{1,2}", clean(q[4]))
            and abs(((q[1] + q[3]) / 2) - yc) < 4
            and q[0] >= w[2] - 3
            and q[0] - w[2] < 35
        ]
        if candidates:
            q = min(candidates, key=lambda z: z[0])
            headers.append({"rn": int(q[4]), "x": w[0], "y": w[1]})
    return list({h["rn"]: h for h in headers}.values())


def _cluster_items(items, value_func, tol=6):
    clusters = []
    for item in sorted(items, key=value_func):
        value = value_func(item)
        if not clusters or abs(value - clusters[-1][0]) > tol:
            clusters.append([value, [item]])
        else:
            clusters[-1][1].append(item)
            clusters[-1][0] = sum(value_func(x) for x in clusters[-1][1]) / len(clusters[-1][1])
    return clusters


def parse_modern_layout_calendar(page, teams):
    words = page.get_text("words") or []
    headers = detect_modern_round_headers(page)
    if len(headers) < 3:
        return [], []

    header_rows = []
    for h in sorted(headers, key=lambda z: (z["y"], z["x"])):
        if not header_rows or abs(h["y"] - header_rows[-1][0]) > 10:
            header_rows.append([h["y"], [h]])
        else:
            header_rows[-1][1].append(h)

    matches = []
    diagnostics = []

    for ri, (hy, group) in enumerate(header_rows):
        group = sorted(group, key=lambda z: z["x"])
        slots = len(group)
        next_hy = header_rows[ri + 1][0] if ri + 1 < len(header_rows) else page.rect.height * .90
        band = [w for w in words if w[1] >= hy - 3 and w[1] < next_hy - 4]

        all_separators = [
            w for w in band
            if clean(w[4]) in {"-", "–", "—"} and w[1] > hy + 15
        ]
        x_clusters = _cluster_items(all_separators, lambda w: (w[0] + w[2]) / 2, 8)

        # Collega ogni intestazione GIORNATA al gruppo di trattini più vicino.
        chosen = []
        used = set()
        for h in group:
            candidates = [
                (abs(c[0] - (h["x"] + 25)), idx, c)
                for idx, c in enumerate(x_clusters)
                if idx not in used and len(c[1]) >= 2
            ]
            if candidates:
                _, idx, c = min(candidates)
                used.add(idx)
                chosen.append(c)

        if len(chosen) != slots:
            chosen = sorted(
                sorted(x_clusters, key=lambda c: len(c[1]), reverse=True)[:slots],
                key=lambda c: c[0],
            )
        else:
            chosen = sorted(chosen, key=lambda c: c[0])

        if len(chosen) != slots:
            diagnostics.append(("xcluster", ri, slots))
            continue

        centers = [c[0] for c in chosen]
        if len(centers) > 1:
            bounds = [max(0, centers[0] - (centers[1] - centers[0]) / 2)]
            bounds += [(a + b) / 2 for a, b in zip(centers, centers[1:])]
            bounds += [min(page.rect.width, centers[-1] + (centers[-1] - centers[-2]) / 2)]
        else:
            bounds = [0, page.rect.width]

        for j, h in enumerate(group):
            x0, x1 = bounds[j], bounds[j + 1]
            sep_words = sorted(chosen[j][1], key=lambda w: (w[1] + w[3]) / 2)
            y_clusters = _cluster_items(sep_words, lambda w: (w[1] + w[3]) / 2, 2.5)
            sep_words = [
                min(cluster[1], key=lambda w: abs(((w[0] + w[2]) / 2) - centers[j]))
                for cluster in y_clusters
            ]
            if not sep_words:
                continue

            ys = [(w[1] + w[3]) / 2 for w in sep_words]
            gaps = [b - a for a, b in zip(ys, ys[1:]) if 4 < b - a < 30]
            game_gap = statistics.median(gaps) if gaps else 10
            y_bounds = [ys[0] - game_gap / 2]
            y_bounds += [(a + b) / 2 for a, b in zip(ys, ys[1:])]
            y_bounds += [ys[-1] + game_gap / 2]

            slot_words = [w for w in band if x0 <= ((w[0] + w[2]) / 2) < x1]

            dates = []
            for w in slot_words:
                if (w[1] + w[3]) / 2 >= y_bounds[0]:
                    continue
                dt = normalize_numeric_date(w[4])
                if dt and dt not in dates:
                    dates.append(dt)

            for k, sep in enumerate(sep_words):
                sep_x = (sep[0] + sep[2]) / 2
                row_words = [
                    w for w in slot_words
                    if y_bounds[k] <= ((w[1] + w[3]) / 2) < y_bounds[k + 1]
                    and clean(w[4]) not in {"-", "–", "—"}
                ]
                left = [w for w in row_words if ((w[0] + w[2]) / 2) < sep_x - 1]
                right = [w for w in row_words if ((w[0] + w[2]) / 2) > sep_x + 1]

                home_raw = clean(" ".join(w[4] for w in sorted(left, key=lambda z: (z[1], z[0]))))
                away_raw = clean(" ".join(w[4] for w in sorted(right, key=lambda z: (z[1], z[0]))))
                if not home_raw or not away_raw:
                    continue

                home = smart_canonical_team(home_raw, teams)
                away = smart_canonical_team(away_raw, teams)

                if home == "RIPOSA" or away == "RIPOSA":
                    continue
                if not home or not away or home == away:
                    diagnostics.append(("unmapped", h["rn"], home_raw, away_raw, home, away))
                    continue
                if not dates:
                    diagnostics.append(("nodate", h["rn"], home_raw, away_raw))
                    continue

                info_home = teams.get(home, TeamInfo(home))
                matches.append(Match(
                    date=dates[0], home=home, away=away,
                    time=info_home.time, locality=info_home.locality, address=info_home.address,
                    round_no=str(h["rn"]),
                ))

                if len(dates) >= 2:
                    info_return = teams.get(away, TeamInfo(away))
                    matches.append(Match(
                        date=dates[1], home=away, away=home,
                        time=info_return.time, locality=info_return.locality, address=info_return.address,
                        round_no=str(h["rn"]),
                    ))

    unique = {}
    for m in matches:
        unique[(m.date, m.home, m.away, m.round_no)] = m
    return list(unique.values()), diagnostics


# ============================================================
# PROGRAMMA GARE / COPPA ITALIA
# ============================================================

def parse_programma_gare_pdf(doc):
    sections = []
    groups = {}
    competition = "COPPA ITALIA ECCELLENZA"

    for page in doc:
        try:
            tables = page.find_tables().tables
        except Exception:
            tables = []
        for table in tables:
            data = table.extract()
            current_group = None
            for row in data:
                vals = [clean(x or "") if x is not None else "" for x in row]
                if not any(vals):
                    continue
                if vals[0].startswith("CAMPIONATO "):
                    competition = re.sub(r"^CAMPIONATO\s+[A-Z]{1,3}\s+", "", vals[0]).strip() or competition
                    continue
                gm = re.match(r"GIRONE\s+(\d+)", vals[0])
                if gm:
                    current_group = gm.group(1)
                    groups.setdefault(current_group, [])
                    continue
                if current_group and len(vals) >= 8:
                    dt = normalize_numeric_date(vals[3])
                    tm = normalize_time(vals[4])
                    if dt and tm and vals[0] and vals[1]:
                        groups[current_group].append(Match(
                            date=dt,
                            home=vals[0],
                            away=vals[1],
                            time=tm,
                            locality=vals[6],
                            address=vals[7],
                            round_no=vals[5],
                        ))

    for group, matches in sorted(groups.items(), key=lambda x: int(x[0])):
        names = sorted({m.home for m in matches} | {m.away for m in matches})
        teams = {name: TeamInfo(name=name) for name in names}
        if matches:
            sections.append(Section(competition, group, teams, matches, "programma_gare_grid"))
    return sections


# ============================================================
# INTESTAZIONE CATEGORIA / GIRONE
# ============================================================

def parse_competition_group(text):
    lines = [clean(x) for x in (text or "").splitlines() if clean(x)]
    group = ""
    for line in lines:
        m = re.search(r"\bGIRONE\s*:?[ ]*([A-Z0-9]+)\b", line)
        if m:
            group = m.group(1)
            break

    competition = ""
    skip = {"FASE AUTUNNALE", "FASE PRIMAVERILE", f"GIRONE {group}" if group else ""}
    for line in lines[:12]:
        if line in skip or line.startswith("STAGIONE") or re.fullmatch(r"20\d{2}/20\d{2}", line):
            continue
        if line.startswith("SOCIETA") or line.startswith("COMITATO") or line.startswith("LOMBARDIA"):
            continue
        if "GIRONE" in line and len(line) < 30:
            continue
        if len(line) >= 4:
            competition = line
            break
    return competition or "CALENDARIO", group


# ============================================================
# PARSER PDF PRINCIPALE
# ============================================================

def parse_pdf(path):
    doc = fitz.open(path)
    all_text = "\n".join(page.get_text("text") for page in doc)

    # Programma Gare è una struttura diversa dai calendari A/R.
    if "PROGRAMMA GARE" in clean(all_text) and "COPPA ITALIA" in clean(all_text):
        sections = parse_programma_gare_pdf(doc)
        if sections:
            return sections

    sections = []
    used = set()

    # Ogni coppia calendario + tabella campi diventa una sezione.
    for pi in range(len(doc) - 1):
        if pi in used:
            continue
        headers = detect_modern_round_headers(doc[pi])
        if len(headers) < 3:
            continue
        teams = parse_team_table(doc[pi + 1])
        if len(teams) < 4:
            continue
        matches, diagnostics = parse_modern_layout_calendar(doc[pi], teams)
        if not matches:
            continue
        competition, group = parse_competition_group(doc[pi + 1].get_text("text"))
        sections.append(Section(competition, group, teams, matches, "modern_layout_v4"))
        used.update([pi, pi + 1])

    return sections


# ============================================================
# DOCX - FALLBACK CLASSICO
# ============================================================

def parse_docx(path):
    """
    Supporto conservativo per vecchi DOCX LND. Il parser cerca calendario classico
    ANDATA/RITORNO e l'elenco campi. Se il documento non corrisponde, restituisce [].
    """
    try:
        from docx import Document
    except Exception:
        return []

    document = Document(path)
    lines = [p.text for p in document.paragraphs if p.text.strip()]
    text = "\n".join(lines)

    # Ricava categoria/girone.
    competition, group = parse_competition_group(text)

    # Estrazione minima delle partite: righe ASCII con " - ".
    raw_matches = []
    current_a = current_r = ""
    round_no = ""
    for line in lines:
        c = clean(line)
        ma = re.search(r"ANDATA:\s*(\d{1,2}/\d{1,2}/\d{2,4})", c)
        mr = re.search(r"RITORNO:\s*(\d{1,2}/\d{1,2}/\d{2,4})", c)
        mg = re.search(r"\b(\d{1,2})\s+G\s*I\s*O\s*R\s*N\s*A\s*T\s*A\b", c)
        if ma:
            current_a = normalize_numeric_date(ma.group(1))
        if mr:
            current_r = normalize_numeric_date(mr.group(1))
        if mg:
            round_no = mg.group(1)

        # I vecchi DOCX possono avere tre riquadri sulla stessa riga: separa i blocchi '| ... |'.
        blocks = re.findall(r"\|([^|]+?)\|", line)
        for block in blocks:
            b = clean(block)
            if " - " not in b or "GIORNATA" in b or "ANDATA" in b or "RITORNO" in b:
                continue
            parts = [x.strip() for x in re.split(r"\s+-\s+", b, maxsplit=1)]
            if len(parts) == 2 and parts[0] and parts[1] and "RIPOSA" not in b:
                if current_a:
                    raw_matches.append((current_a, parts[0], parts[1], round_no))
                if current_r:
                    raw_matches.append((current_r, parts[1], parts[0], round_no))

    if not raw_matches:
        return []

    names = sorted({h for _, h, _, _ in raw_matches} | {a for _, _, a, _ in raw_matches})
    teams = {n: TeamInfo(n) for n in names}
    matches = [Match(d, h, a, round_no=r) for d, h, a, r in raw_matches]
    return [Section(competition, group, teams, matches, "docx_classic")]


# ============================================================
# DATE E GIORNO SOCIETA'
# ============================================================

def adjusted_match_date(section, match):
    """
    A./R. = data ufficiale della delegazione.
    - Giorno vuoto: data invariata.
    - Sabato/Domenica: allinea solo se la data ufficiale cade nel weekend.
    - Turni infrasettimanali: invariati.
    """
    try:
        dt = datetime.strptime(match.date, "%d/%m/%Y")
    except Exception:
        return match.date

    info = section.teams.get(match.home)
    declared = clean(info.day) if info else ""
    declared = declared.replace("Ì", "I").replace("Í", "I")
    if not declared:
        return dt.strftime("%d/%m/%Y")

    wd = dt.weekday()  # lun=0, sab=5, dom=6
    if wd not in (5, 6):
        return dt.strftime("%d/%m/%Y")

    if declared == "SABATO" and wd == 6:
        dt -= timedelta(days=1)
    elif declared == "DOMENICA" and wd == 5:
        dt += timedelta(days=1)

    return dt.strftime("%d/%m/%Y")


def adjusted_sort_key(section, match):
    return (sort_date_value(adjusted_match_date(section, match)), match.time, match.home, match.away)


# ============================================================
# EXCEL
# ============================================================

def excel_date_value(value):
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    return value


def excel_time_value(value):
    value = normalize_time(value)
    if not value:
        return ""
    try:
        return datetime.strptime(value, "%H:%M").time()
    except ValueError:
        return value


def create_excel_for_team(section, selected_team):
    selected = [m for m in section.matches if m.home == selected_team or m.away == selected_team]
    selected = sorted(selected, key=lambda m: adjusted_sort_key(section, m))

    wb = Workbook()
    ws = wb.active
    ws.title = "Calendario"
    ws.append(["Data", "Ora", "Tipo", "Squadra casa", "Squadra ospite", "Indirizzo"])

    for m in selected:
        ws.append([
            excel_date_value(adjusted_match_date(section, m)),
            excel_time_value(m.time),
            "CAMPIONATO",
            m.home,
            m.away,
            indirizzo_excel(m),
        ])
        row = ws.max_row
        ws.cell(row=row, column=1).number_format = "dd/mm/yy"
        ws.cell(row=row, column=2).number_format = "hh:mm"
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")

    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")

    widths = [14, 10, 16, 32, 32, 52]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), selected


def safe_filename(value):
    s = clean(value)
    s = re.sub(r"[^A-Z0-9]+", "_", s).strip("_")
    return s or "SQUADRA"


# ============================================================
# CACHE / DIAGNOSTICA
# ============================================================

@st.cache_data(show_spinner=False)
def analyze_upload_v4(file_bytes, filename, parser_version):
    suffix = Path(filename).suffix.lower()
    if suffix not in {".pdf", ".docx"}:
        raise ValueError("Sono supportati file PDF e DOCX.")

    tmp = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
            f.write(file_bytes)
            tmp = f.name
        return parse_docx(tmp) if suffix == ".docx" else parse_pdf(tmp)
    finally:
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


def diagnose_pdf_bytes(file_bytes):
    info = {"pagine": 0, "giornate_p1": 0, "squadre_p2": 0, "gare_layout": 0, "non_mappate": 0}
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
            f.write(file_bytes)
            tmp = f.name
        doc = fitz.open(tmp)
        info["pagine"] = len(doc)
        if len(doc) >= 1:
            info["giornate_p1"] = len(detect_modern_round_headers(doc[0]))
        if len(doc) >= 2:
            teams = parse_team_table(doc[1])
            info["squadre_p2"] = len(teams)
            if teams:
                matches, diag = parse_modern_layout_calendar(doc[0], teams)
                info["gare_layout"] = len(matches)
                info["non_mappate"] = sum(1 for x in diag if x and x[0] == "unmapped")
        return info
    except Exception as exc:
        info["errore"] = str(exc)
        return info
    finally:
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


# ============================================================
# STREAMLIT
# ============================================================

st.set_page_config(page_title="Calendario → Excel", page_icon="⚽", layout="centered")
st.title("⚽ Calendario → Excel")
st.caption(f"Versione app: {PARSER_VERSION}")
st.write("Carica un calendario LND/FIGC, scegli il girone e la squadra, quindi scarica l’Excel.")

uploaded = st.file_uploader("1. Carica il calendario", type=["pdf", "docx"])

if uploaded is not None:
    with st.spinner("Analisi del calendario in corso…"):
        try:
            sections = analyze_upload_v4(uploaded.getvalue(), uploaded.name, PARSER_VERSION)
        except Exception as exc:
            st.error(f"Errore durante la lettura del file: {exc}")
            st.stop()

    if not sections:
        st.error("Non sono riuscito a riconoscere partite nel file.")
        if Path(uploaded.name).suffix.lower() == ".pdf":
            diag = diagnose_pdf_bytes(uploaded.getvalue())
            with st.expander("Diagnostica PDF", expanded=True):
                st.write(f"Pagine lette: **{diag.get('pagine', 0)}**")
                st.write(f"Intestazioni GIORNATA rilevate nella pagina 1: **{diag.get('giornate_p1', 0)}**")
                st.write(f"Squadre rilevate nella tabella della pagina 2: **{diag.get('squadre_p2', 0)}**")
                st.write(f"Righe gara ricostruite dal layout: **{diag.get('gare_layout', 0)}**")
                st.write(f"Righe non associate a una squadra: **{diag.get('non_mappate', 0)}**")
                if diag.get("errore"):
                    st.code(diag["errore"])
        st.stop()

    st.success(f"Analisi completata: {len(sections)} sezione/i riconosciuta/e.")

    labels = []
    for idx, section in enumerate(sections, 1):
        label = section.label or f"Sezione {idx}"
        if label in labels:
            label = f"{label} ({idx})"
        labels.append(label)

    if len(sections) > 1:
        chosen_label = st.selectbox("2. Seleziona categoria / girone", labels)
        section = sections[labels.index(chosen_label)]
    else:
        section = sections[0]
        st.info(f"Categoria/Girone: {section.label}")

    teams = sorted({m.home for m in section.matches} | {m.away for m in section.matches})
    if not teams:
        st.error("Nessuna squadra riconosciuta nella sezione selezionata.")
        st.stop()

    selected_team = st.selectbox("3. Per quale squadra vuoi l'estrapolazione?", teams)
    team_matches = [m for m in section.matches if m.home == selected_team or m.away == selected_team]
    st.write(f"Partite trovate per **{selected_team}**: **{len(team_matches)}**")

    with st.expander("Dettagli analisi"):
        st.write(f"Versione parser: `{PARSER_VERSION}`")
        st.write(f"Formato riconosciuto: `{section.source_format}`")
        st.write(f"Squadre nel girone: **{len(teams)}**")
        st.write(f"Partite complessive lette: **{len(section.matches)}**")

    if st.button("4. Genera Excel", type="primary", use_container_width=True):
        excel_bytes, extracted = create_excel_for_team(section, selected_team)
        st.session_state["excel_bytes"] = excel_bytes
        st.session_state["excel_name"] = f"Calendario_{safe_filename(selected_team)}.xlsx"
        st.session_state["excel_count"] = len(extracted)
        st.session_state["excel_key"] = (uploaded.name, section.label, selected_team, PARSER_VERSION)

    current_key = (uploaded.name, section.label, selected_team, PARSER_VERSION)
    if st.session_state.get("excel_key") == current_key and "excel_bytes" in st.session_state:
        st.success(f"Excel pronto: {st.session_state.get('excel_count', 0)} partite estratte.")
        st.download_button(
            "⬇️ Scarica Excel",
            data=st.session_state["excel_bytes"],
            file_name=st.session_state["excel_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.divider()
st.caption("Colonne Excel: Data | Ora | CAMPIONATO | Squadra casa | Squadra ospite | Indirizzo - Paese")
