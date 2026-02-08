# Extraktor e-mailů

Tento Python skript zpracovává soubory `.eml` a `.msg` za účelem extrakce e-mailových adres a jejich uložení do CSV souborů. Je navržen tak, aby efektivně zvládal velké množství souborů s možností nastavení počtu souborů ke zpracování.

## Funkce

### Základní funkcionalita

- **Extrakce e-mailových adres** ze souborů `.eml` (standardní e-mailový formát) a `.msg` (Microsoft Outlook formát)
- **Validace e-mailových adres** pomocí regulárních výrazů pro zajištění správnosti formátu
- **Identifikace souborů** pomocí ID získaného z názvu souboru (část před "---")
- **Sledování času zpracování** s výpisem celkové doby běhu skriptu

### Pokročilé zpracování

- **Multipart e-maily**: Zpracovává složité e-mailové struktury s více částmi (text/plain, text/html)
- **Přeposlané zprávy**: Vyhledává další e-mailové adresy v těle zprávy pomocí vzoru "From:"
- **Duplicitní kontrola**: Automaticky odstraňuje duplicitní e-mailové adresy v rámci jednoho souboru
- **Robustní zpracování chyb**: Pokračuje ve zpracování i při chybách jednotlivých souborů

### Výstupní soubory

Skript vytváří dva CSV soubory s oddělovačem středník (;):

1. **`vystup_all.csv`**: Obsahuje všechny extrahované e-mailové adresy
   - Struktura: `id_souboru;email1;email2;email3;...`
   - Zachovává všechny nalezené e-mailové adresy bez filtrace

2. **`vystup_bez_ct_email.csv`**: Filtrovaná verze bez interních e-mailů České televize
   - Odstraňuje e-mailové adresy končící na `@ceskatelevize.cz`
   - Zachovává pozice ostatních e-mailů (CT e-maily jsou nahrazeny prázdnými řetězci)
   - Užitečné pro analýzu externí komunikace

## Požadavky

- Python 3.x
- Požadované Python moduly:
  - `os` (standardní knihovna)
  - `csv` (standardní knihovna)
  - `re` (standardní knihovna)
  - `time` (standardní knihovna)
  - `email` (standardní knihovna)
  - `extract_msg` (externí knihovna - vyžaduje instalaci)

### Instalace externích závislostí

```bash
pip install extract-msg
```

## Konfigurace

Ve skriptu můžete upravit následující parametry:

- `slozka`: Cesta ke složce s `.eml` a `.msg` soubory
- `max_soubory`: Maximální počet souborů ke zpracování (výchozí: 25 000)
- `vystup_csv`: Název výstupního CSV souboru (nepoužívá se v aktuální verzi)

## Použití

1. Umístěte své `.eml` a `.msg` soubory do složky specifikované proměnnou `slozka` ve skriptu
2. Ujistěte se, že máte nainstalované všechny požadované závislosti
3. Spusťte skript:

   ```bash
   python email_extractor.py
   ```

4. Skript zobrazí průběh zpracování a na konci vytiskne dobu běhu
5. Výstupní CSV soubory budou vytvořeny ve stejné složce jako skript

## Výstup

Skript během běhu vypisuje:

- Identifikátor souboru a nalezené e-mailové adresy pro každý zpracovaný soubor
- Informace o uložení výstupních souborů
- Celkovou dobu běhu skriptu v sekundách

## Technické detaily

### Zpracování .eml souborů

- Používá `email.parser.BytesParser` s `email.policy.default`
- Podporuje multipart zprávy s různými obsahovými typy
- Extrahuje e-maily z hlavičky "From" a z těla zprávy

### Zpracování .msg souborů

- Používá knihovnu `extract_msg` pro čtení Outlook souborů
- Extrahuje e-maily z vlastnosti `sender` a z těla zprávy

### Validace e-mailů

- Regulární výraz: `^[^@\s]+@[^@\s]+\.[a-zA-Z0-9]+$`
- Kontroluje základní formát e-mailové adresy

### Zpracování chyb

- Pokračuje ve zpracování i při chybách jednotlivých souborů
- Vypisuje chybové zprávy s informací o problematickém souboru