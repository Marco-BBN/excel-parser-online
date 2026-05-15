import io

import openpyxl
import streamlit as st

from parsers import (
    DEFAULT_SKIP_SHEETS,
    build_rows,
    cards_to_fragment,
    check_issues,
    parse_cards,
    parse_tables,
    tables_to_fragment,
)

st.set_page_config(page_title="Excel → HTML", page_icon="📊", layout="wide")
st.title("📊 Excel → HTML")

PREVIEW_CSS = """
<style>
  body { font-family: sans-serif; font-size: 14px; margin: 12px; color: #222; }
  h4 { margin: 20px 0 6px; font-size: 15px; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
  th, td { border: 1px solid #ddd; padding: 6px 10px; text-align: left; vertical-align: top; }
  th { background: #f0f0f0; font-weight: 600; }
  tr:nth-child(even) td { background: #fafafa; }
</style>
"""

# ---------------------------------------------------------------------------
# Sidebar — nastavení
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("Nastavení")
    mode = st.radio("Typ parseru", ["Tabulky (horizontální)", "Kartičky (vertikální)"])
    skip_input = st.text_input(
        "Přeskočit záložky (oddělené čárkou)",
        value=", ".join(DEFAULT_SKIP_SHEETS),
    )
    skip = {s.strip().lower() for s in skip_input.split(",") if s.strip()}

# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------

uploaded = st.file_uploader("Nahraj Excel soubor (.xlsx)", type=["xlsx"])

if not uploaded:
    st.info("Nahraj .xlsx soubor pro začátek.")
    st.stop()

wb = openpyxl.load_workbook(io.BytesIO(uploaded.read()))
sheet_names = [s for s in wb.sheetnames if s.lower() not in skip]

if not sheet_names:
    st.warning("Žádné záložky k zobrazení — všechny jsou v seznamu přeskočených.")
    st.stop()

# ---------------------------------------------------------------------------
# Záložky (jedna per sheet)
# ---------------------------------------------------------------------------

tabs = st.tabs(sheet_names)

for tab, sheet_name in zip(tabs, sheet_names):
    with tab:
        sheet = wb[sheet_name]
        is_cards = mode.startswith("Kartičky")

        if is_cards:
            items = parse_cards(sheet)
            fragment = cards_to_fragment(items) if items else ""
            issues = check_issues(sheet_name, items) if items else []
            count_label = f"{len(items)} produktů"
        else:
            rows = build_rows(sheet)
            items = parse_tables(rows)
            fragment = tables_to_fragment(items) if items else ""
            issues = []
            count_label = f"{len(items)} tabulek"

        if not items:
            st.warning(f"Záložka '{sheet_name}': žádný obsah nenalezen.")
            continue

        st.caption(count_label)

        if issues:
            with st.expander(f"⚠ Upozornění ({len(issues)})", expanded=False):
                for issue in issues:
                    st.warning(issue)

        # --- Náhled (lze označit a CTRL+C do WYSIWYG) ---
        row_count = sum(len(t["rows"]) for t in items) if not is_cards else sum(len(p["params"]) for p in items)
        preview_height = max(300, min(800, 60 + row_count * 28 + len(items) * 50))

        st.components.v1.html(PREVIEW_CSS + fragment, height=preview_height, scrolling=True)

        # --- HTML ke zkopírování do WordPress HTML editoru ---
        with st.expander("📋 HTML kód (pro WordPress HTML editor)"):
            st.code(fragment, language="html")
