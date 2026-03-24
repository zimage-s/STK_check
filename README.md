# STK Check

Kontrola platnosti STK vozidel z XLSX souboru přes [kontrolatachaku.cz](https://www.kontrolatachaku.cz) (data z Ministerstva dopravy ČR).

## Požadavky

```bash
pip3 install requests beautifulsoup4 openpyxl
```

## Použití

```bash
python3 stk_check.py stahni              # stáhne STK data z webu → stk_data.json
python3 stk_check.py xlsx                # z stk_data.json → STK_vysledky.xlsx
python3 stk_check.py stahni xlsx         # obojí najednou
python3 stk_check.py stahni --vin VIN    # stáhne/doplní jen jedno auto podle VIN
```

## Vstup

XLSX soubor s listem `List1`, kde:
- sloupec **H** = značka
- sloupec **I** = model
- sloupec **K** = registrační značka (RZ)
- sloupec **L** = VIN
- sloupec **N** = rok výroby

## Výstup

**`STK_vysledky.xlsx`** se dvěma listy:

1. **STK Kontrola** – přehled všech vozidel s posledním STK, příštím termínem a barevným stavem (OK / blíží se / po lhůtě)
2. **Historie STK** – kompletní záznamy všech prohlídek

## Zdroj dat

Oficiální data z [datové kostky Ministerstva dopravy ČR](https://www.kontrolatachaku.cz) (licence ODVS). Záznamy existují přibližně od roku 2008.
