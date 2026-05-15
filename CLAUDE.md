# CLAUDE.md

Tento soubor poskytuje návod pro Claude Code při práci s tímto repozitářem.

## Přehled projektu

Streamlit webová aplikace pro konverzi Excel souborů (`.xlsx`) do HTML fragmentů vhodných pro vložení do WordPress (nebo jiného CMS) přes WYSIWYG nebo HTML editor.

**Nasazení:** Streamlit Community Cloud — [share.streamlit.io](https://share.streamlit.io)
**GitHub repo:** `Marco-BBN/excel-parser-online`
**GitHub účet:** `ondrejvrtel` (přihlašovací email: ondrej.vrtel@marco.eu)

## Struktura souborů

```
excel-parser-online/
├── app.py            # Streamlit UI — jediný vstupní bod aplikace
├── parsers.py        # Veškerá logika parsování (bez CLI)
├── requirements.txt  # Závislosti: streamlit, openpyxl
└── CLAUDE.md
```

Původní CLI skripty (ze složky `../excel-parser/`) jsou zachovány samostatně a nejsou součástí této aplikace. Logika byla extrahována do `parsers.py`.

## Architektura

### `parsers.py`

Obsahuje dvě sady funkcí:

**Tabulky (horizontální)** — odpovídá `excel_to_html.py`:
- `build_rows(sheet)` — rozepíše sloučené buňky
- `parse_tables(rows)` — detekuje tabulky podle `HEADER_TRIGGERS`
- `tables_to_fragment(tables)` — generuje HTML fragment (bez `<html>/<body>`)

**Kartičky (vertikální)** — odpovídá `excel_to_html_cards.py`:
- `parse_cards(sheet)` — detekuje produkty jako vertikální karty
- `cards_to_fragment(products)` — generuje HTML fragment
- `check_issues(sheet_name, products)` — validace (chybějící/duplicitní kódy)

Výstup jsou vždy **HTML fragmenty** (ne celé stránky) — pouze `<h4>` + `<table class="basic-table">` bloky, připravené pro přímé vložení do WordPressu.

### `app.py`

- Upload `.xlsx` přes `st.file_uploader`
- Sidebar: přepínač Tabulky / Kartičky + pole pro přeskočené záložky (výchozí: `seznam`)
- Každá záložka Excelu = jeden Streamlit tab (`st.tabs`)
- **Náhled** v `st.components.v1.html()` — lze označit a CTRL+C do WYSIWYG editoru
- **HTML kód** v `st.code()` — tlačítko kopírovat pro WordPress HTML editor

## Pravidla pro strukturu Excelu

### Tabulky (horizontální)
- Hlavičkový řádek začíná slovem: `Typ`, `Příslušenství`, `Čidlo:`, `Použití:`
- Titulek tabulky je v řádku těsně nad hlavičkou
- Prázdný řádek odděluje tabulky
- Řádky začínající `Krytí:`, `Doporučená:`, `Napětí:` jsou přeskočeny (poznámky)
- Záložka `seznam` je přeskočena ve výchozím nastavení

### Kartičky (vertikální)
- Název produktu = jediná neprázdná buňka v prvním sloupci
- Řádek `Typ` (samotný) je oddělovač — přeskočen
- Parametry: první sloupec = název, druhý sloupec = hodnota
- Prázdný řádek odděluje produkty

## Nasazení a provoz

**Deployment** probíhá automaticky při každém `git push` na větev `main`.

```bash
# Lokální vývoj
pip install streamlit openpyxl
streamlit run app.py

# Push = automatický redeploy na Streamlit Cloud
git add .
git commit -m "popis změny"
git push
```

**SSH klíč** pro push je nastaven na GitHub účtu `ondrejvrtel` (soubor `~/.ssh/id_ed25519`).

## Rozšíření a úpravy

**Přidat nový trigger pro detekci hlavičky tabulky:**
→ upravte `HEADER_TRIGGERS` tuple v `parsers.py`

**Přidat nový prefix pro přeskočení řádku:**
→ upravte `SKIP_PREFIXES` tuple v `parsers.py`

**Změnit výchozí přeskočené záložky:**
→ upravte `DEFAULT_SKIP_SHEETS` set v `parsers.py`

**Upravit CSS náhledu** (styl tabulek v preview, ne ve WordPressu):
→ upravte `PREVIEW_CSS` string v `app.py`

**Přidat třetí typ parseru:**
1. Přidejte parse + fragment funkce do `parsers.py`
2. Přidejte volbu do `st.radio` v `app.py`
3. Přidejte větev do podmínky `if/elif` v hlavní smyčce přes záložky

## Workflow uživatele

1. Nahrát `.xlsx` soubor
2. Vybrat typ parseru (Tabulky / Kartičky)
3. Procházet záložky (= záložky v Excelu)
4. **Pro WYSIWYG editor:** označit tabulku v náhledu → CTRL+C → vložit do WordPress vizuálního editoru
5. **Pro HTML editor:** rozbalit "📋 HTML kód" → kliknout kopírovat → vložit do WordPress HTML editoru

## Jazyk

Aplikace i komentáře jsou v češtině. WordPress třídy (`basic-table`) zachovat beze změny — jsou definovány v šabloně WordPressu.
