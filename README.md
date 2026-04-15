# STK Check

Kontrola platnosti STK vozidel z XLSX souboru.

**Primární zdroj:** [API Ministerstva dopravy ČR](https://www.dataovozidlech.cz/) (real-time data z registru vozidel)
**Fallback:** [kontrolatachaku.cz](https://www.kontrolatachaku.cz) (pro VINy, které API nezná — typicky krátké VINy veteránů)

## Požadavky

```bash
pip3 install requests beautifulsoup4 openpyxl
```

API klíč z [dataovozidlech.cz/registraceApi](https://dataovozidlech.cz/registraceApi) (zdarma, registrace přes email).
Klíč uložte do `config.json`:

```json
{ "md_api_key": "VÁŠ_KLÍČ" }
```

## Použití

```bash
python3 stk_check.py stahni              # stáhne STK data → stk_data.json
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
2. **Historie STK** – kompletní záznamy všech prohlídek (pouze z webu, API historii neposkytuje)

## Zdroje dat

| Zdroj | Typ dat | VINy | Aktuálnost |
|---|---|---|---|
| API MD ČR (`api.dataovozidlech.cz`) | registr vozidel + datum příští STK | standardní (11-17 zn.) | real-time |
| kontrolatachaku.cz | historie STK prohlídek | všechny vč. krátkých | zpoždění týdny/měsíce |

Skript automaticky zkouší API jako první. Pokud VIN nenajde, použije web jako fallback.
