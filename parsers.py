"""
Parsing logic extracted from excel_to_html.py and excel_to_html_cards.py.
Produces HTML fragments (no full page wrapper) suitable for embedding.
"""

import datetime
import re


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def escape_html(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def clean_str(v: object) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime.time):
        return f"{v.hour} A"
    if isinstance(v, datetime.datetime):
        return v.strftime("%d.%m.%Y")
    s = str(v).replace("\xa0", " ").strip()
    s = re.sub(r"min,\s*", "min. ", s)
    s = re.sub(r"max,\s*", "max. ", s)
    return s


def format_val(v: object) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v).replace(".", ",")
    return clean_str(v)


# ---------------------------------------------------------------------------
# Tables (horizontal) — excel_to_html.py logic
# ---------------------------------------------------------------------------

HEADER_TRIGGERS = ("typ", "příslušenství", "čidlo:", "použití:")
SKIP_PREFIXES = ("krytí", "doporučená", "napětí")
DEFAULT_SKIP_SHEETS = {"seznam"}


def resolve_merged(sheet) -> dict:
    resolved = {}
    for mr in sheet.merged_cells.ranges:
        val = sheet.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                resolved[(r, c)] = val
    return resolved


def build_rows(sheet) -> list:
    resolved = resolve_merged(sheet)
    merge_non_top = set()
    for mr in sheet.merged_cells.ranges:
        for r in range(mr.min_row + 1, mr.max_row + 1):
            merge_non_top.add(r)

    def should_skip(r):
        if r not in merge_non_top:
            return False
        for c in range(1, sheet.max_column + 1):
            if sheet.cell(r, c).value not in (None, ""):
                return False
        return True

    rows = []
    for r in range(1, sheet.max_row + 1):
        if should_skip(r):
            continue
        row = [resolved.get((r, c), sheet.cell(r, c).value) for c in range(1, sheet.max_column + 1)]
        rows.append(row)
    return rows


def parse_tables(rows: list) -> list:
    tables = []
    i = 0
    while i < len(rows):
        row = rows[i]
        non_empty = [(j, v) for j, v in enumerate(row) if v is not None and v != ""]
        if not non_empty:
            i += 1
            continue
        first_val = clean_str(non_empty[0][1])
        if first_val.lower() in HEADER_TRIGGERS:
            headers = [clean_str(v) for _, v in non_empty]
            col_indices = [j for j, _ in non_empty]
            title = ""
            for back in range(i - 1, -1, -1):
                bnon = [(j, v) for j, v in enumerate(rows[back]) if v is not None and v != ""]
                if bnon:
                    candidate = clean_str(bnon[0][1])
                    if not any(candidate.lower().startswith(p) for p in SKIP_PREFIXES):
                        title = candidate
                    break
            data_rows = []
            j = i + 1
            while j < len(rows):
                drow = rows[j]
                dnon = [(jj, v) for jj, v in enumerate(drow) if v is not None and v != ""]
                if not dnon:
                    break
                fd = clean_str(dnon[0][1])
                if fd.lower() in HEADER_TRIGGERS:
                    break
                if any(fd.lower().startswith(p) for p in SKIP_PREFIXES):
                    j += 1
                    continue
                cells = [drow[jj] if jj < len(drow) else None for jj in col_indices]
                if any(v is not None and v != "" for v in cells):
                    data_rows.append(cells)
                j += 1
            if data_rows:
                tables.append({"title": title, "headers": headers, "rows": data_rows})
            i = j
        else:
            i += 1
    return tables


def tables_to_fragment(tables: list) -> str:
    lines = []
    for t in tables:
        if t["title"]:
            lines.append(f'<h4>{escape_html(t["title"])}</h4>')
        lines.append('<table class="basic-table">')
        lines.append("<thead><tr>")
        for h in t["headers"]:
            lines.append(f"<th>{escape_html(h)}</th>")
        lines.append("</tr></thead>")
        lines.append("<tbody>")
        for row in t["rows"]:
            lines.append("<tr>")
            for cell in row:
                lines.append(f"<td>{escape_html(format_val(cell))}</td>")
            lines.append("</tr>")
        lines.append("</tbody></table>")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Cards (vertical) — excel_to_html_cards.py logic
# ---------------------------------------------------------------------------

KOD_PARAM = "Kód"


def parse_cards(sheet) -> list:
    all_rows = list(sheet.iter_rows(values_only=True))
    products = []
    i = 0
    while i < len(all_rows):
        row = all_rows[i]
        non_empty = [(j, c) for j, c in enumerate(row) if c is not None and c != ""]
        if not non_empty:
            i += 1
            continue
        first_val = clean_str(non_empty[0][1])
        is_title = (len(non_empty) == 1 and non_empty[0][0] == 0 and first_val.lower() != "typ")
        if is_title:
            title = first_val
            i += 1
            while i < len(all_rows):
                nr = [(j, c) for j, c in enumerate(all_rows[i]) if c is not None and c != ""]
                if nr:
                    if clean_str(nr[0][1]).lower() == "typ" and len(nr) == 1:
                        i += 1
                    break
                i += 1
            params = []
            while i < len(all_rows):
                prow = all_rows[i]
                pnon = [(j, c) for j, c in enumerate(prow) if c is not None and c != ""]
                if not pnon:
                    break
                pf = clean_str(pnon[0][1])
                if len(pnon) == 1 and pnon[0][0] == 0 and pf.lower() != "typ":
                    break
                key = pf.rstrip(":").strip()
                val = clean_str(pnon[1][1]) if len(pnon) > 1 else ""
                if key:
                    params.append((key, val))
                i += 1
            if params:
                products.append({"title": title, "params": params})
        else:
            i += 1
    return products


def cards_to_fragment(products: list) -> str:
    lines = []
    for p in products:
        lines.append('<table class="basic-table">')
        lines.append("<thead><tr>")
        lines.append(f'<th colspan="2">{escape_html(p["title"])}</th>')
        lines.append("</tr></thead>")
        lines.append("<tbody>")
        for key, val in p["params"]:
            lines.append(f"<tr><td>{escape_html(key)}</td><td>{escape_html(val)}</td></tr>")
        lines.append("</tbody></table>")
    return "\n".join(lines)


def check_issues(sheet_name: str, products: list) -> list:
    issues = []
    seen_codes = {}
    for p in products:
        params_dict = dict(p["params"])
        title = p["title"]
        kod = params_dict.get(KOD_PARAM, "").strip()
        if not kod:
            issues.append(f"Chybí '{KOD_PARAM}' u produktu '{title}'")
        elif kod in seen_codes:
            issues.append(f"Duplicitní kód {kod} — '{seen_codes[kod]}' a '{title}'")
        else:
            seen_codes[kod] = title
    return issues
