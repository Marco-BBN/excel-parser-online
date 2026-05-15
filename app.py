import io

import openpyxl
import streamlit as st

from parsers import (
    DEFAULT_SKIP_SHEETS,
    build_rows,
    card_to_fragment,
    cards_to_fragment,
    check_issues,
    parse_cards,
    parse_tables,
    table_to_fragment,
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
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("Nastavení")
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

        rows = build_rows(sheet)
        tables = parse_tables(rows)
        cards = parse_cards(sheet)

        if not tables and not cards:
            st.warning(f"Záložka '{sheet_name}': žádný obsah nenalezen.")
            continue

        # Souhrnný náhled pro select+copy do WYSIWYG
        all_html = tables_to_fragment(tables) + ("\n" if tables and cards else "") + cards_to_fragment(cards)
        row_count = sum(len(t["rows"]) for t in tables) + sum(len(p["params"]) for p in cards)
        preview_height = max(300, min(900, 60 + row_count * 28 + (len(tables) + len(cards)) * 50))

        with st.expander("🔍 Náhled celého listu (označit → CTRL+C do WYSIWYG)", expanded=True):
            st.components.v1.html(PREVIEW_CSS + all_html, height=preview_height, scrolling=True)

        st.divider()

        # Upozornění karet
        if cards:
            issues = check_issues(sheet_name, cards)
            if issues:
                with st.expander(f"⚠ Upozornění ({len(issues)})"):
                    for issue in issues:
                        st.warning(issue)

        # Per-tabulka copy sekce
        total = len(tables) + len(cards)
        st.caption(f"{total} {'blok' if total == 1 else 'bloky' if 2 <= total <= 4 else 'bloků'} — HTML kód ke zkopírování:")

        for t in tables:
            label = t["title"] or "(bez názvu)"
            with st.expander(f"📋 {label}"):
                st.code(table_to_fragment(t), language="html")

        for p in cards:
            with st.expander(f"📋 {p['title']}"):
                st.code(card_to_fragment(p), language="html")
