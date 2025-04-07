# Extraktor e-mailů

Tento Python skript zpracovává soubory `.eml` a `.msg` za účelem extrakce e-mailových adres a jejich uložení do CSV souborů. Je navržen tak, aby efektivně zvládal velké množství souborů.

## Funkce

- Extrahuje e-mailové adresy ze souborů `.eml` a `.msg`.
- Validuje e-mailové adresy pro zajištění správnosti.
- Vytváří dva CSV soubory:
  - `vystup_all.csv`: Obsahuje všechny extrahované e-mailové adresy.
  - `vystup_bez_ct_email.csv`: Filtrovány jsou e-mailové adresy končící na `@ceskatelevize.cz`.
- Zpracovává multipart těla e-mailů a vložené e-mailové adresy.

## Požadavky

- Python 3.x
- Požadované Python moduly:
  - `os`
  - `csv`
  - `re`
  - `time`
  - `extract_msg`
  - `email`

## Použití

1. Umístěte své `.eml` a `.msg` soubory do složky specifikované proměnnou `slozka` ve skriptu.
2. Spusťte skript:
   ```bash
   python email_extractor.py
   ```