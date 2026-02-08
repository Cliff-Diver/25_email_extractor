import os
import csv
import re
import time  # Import time module
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
import extract_msg
from email.utils import parseaddr
from datetime import datetime

# --- Nastavení ---
slozka = r"C:\Users\sd232665\OneDrive - Česká televize\Plocha\EDM Doručené maily\test"
vystup_csv = "vystup.csv" # Výstup z skriptu se ukládá do aktuálního pracovního adresáře skriptu
max_soubory = 28000

# --- Pomocné funkce ---
def ziskat_id(nazev):
    return nazev.split("---")[0] if "---" in nazev else nazev

def extrahuj_email(text):
    """Vrací čistou e-mailovou adresu z 'Jméno <email>' nebo jen 'email'"""
    return parseaddr(text)[1]

def je_validni_email(email):
    """Kontroluje, zda je e-mailová adresa validní."""
    return bool(re.match(r'^[^@\s]+@[^@\s]+\.[a-zA-Z0-9]+$', email))

def zpracuj_eml(cesta):
    try:
        with open(cesta, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        # 1) Date:
        date_hdr = msg.get('Date')
        recv_dt = None
        if date_hdr:
            try:
                recv_dt = parsedate_to_datetime(date_hdr)
            except Exception as e:
                print(f"[EML] Chyba parsování Date: {e}")

        # 2) Fallback: všechny Received headers
        if not recv_dt:
            rec_hdrs = msg.get_all('Received', [])
            for rec_hdr in rec_hdrs:
                # Zkusíme různé formáty
                parts = rec_hdr.split(';')
                if len(parts) >= 2:
                    date_part = parts[-1].strip()
                    try:
                        recv_dt = parsedate_to_datetime(date_part)
                        break
                    except Exception:
                        continue

        # 3) Fallback na file modification time
        if not recv_dt:
            try:
                import os
                mtime = os.path.getmtime(cesta)
                recv_dt = datetime.fromtimestamp(mtime)
                print(f"[EML] Použit file mtime pro {os.path.basename(cesta)}")
            except Exception:
                pass

        # 4) E-maily (vaše dosavadní extrakce)
        emaily = []
        hlavni = extrahuj_email(msg.get('From',''))
        if hlavni and je_validni_email(hlavni):
            emaily.append(hlavni)
        if msg.is_multipart():
            obsah = "\n".join(p.get_content() for p in msg.walk()
                              if p.get_content_type() in ('text/plain','text/html'))
        else:
            obsah = msg.get_content()
        for m in re.findall(r'^From:\s+(.+?@.+?)$', obsah,
                            flags=re.MULTILINE|re.IGNORECASE):
            e = extrahuj_email(m)
            if e and je_validni_email(e) and e not in emaily:
                emaily.append(e)

        return recv_dt, emaily

    except Exception as e:
        print(f"[EML] Chyba při zpracování {cesta}: {e}")
        return None, []


def zpracuj_msg(cesta_k_souboru):
    msg = None  # Přidat inicializaci
    try:
        msg = extract_msg.Message(cesta_k_souboru)
        recv_dt = None
        
        # Rozšířené ladění
        print(f"[MSG DEBUG] Zpracovávám: {os.path.basename(cesta_k_souboru)}")
        
        # 1) Raw headers
        try:
            raw_hdr = msg.header() if callable(msg.header) else msg.header
            if isinstance(raw_hdr, str):
                hdr_bytes = raw_hdr.encode('utf-8', errors='ignore')
                parsed = BytesParser(policy=policy.default).parsebytes(hdr_bytes)
                date_hdr = parsed.get('Date')
                if date_hdr:
                    recv_dt = parsedate_to_datetime(date_hdr)
        except Exception as e:
            print(f"[MSG] Chyba raw headers: {e}")

        # 2) msg.date attribute - OPRAVA
        if not recv_dt and hasattr(msg, 'date') and msg.date:
            try:
                # Pokud je msg.date už datetime objekt, použij přímo
                if isinstance(msg.date, datetime):
                    recv_dt = msg.date
                else:
                    # Pokud je string, parsuj
                    recv_dt = parsedate_to_datetime(msg.date)
            except Exception as e:
                print(f"[MSG] Chyba msg.date: {e}")

        # 3) File modification time fallback
        if not recv_dt:
            try:
                mtime = os.path.getmtime(cesta_k_souboru)
                recv_dt = datetime.fromtimestamp(mtime)
                print(f"[MSG] Použit file mtime pro {os.path.basename(cesta_k_souboru)}")
            except Exception:
                pass

        # 4) Extrakce e-mailů
        emaily = []
        hlavni = extrahuj_email(msg.sender or '')
        if hlavni and je_validni_email(hlavni):
            emaily.append(hlavni)

        for nalezeny in re.findall(r'^From:\s+(.+?@.+?)$', msg.body or '',
                                   flags=re.MULTILINE|re.IGNORECASE):
            e = extrahuj_email(nalezeny)
            if e and je_validni_email(e) and e not in emaily:
                emaily.append(e)

        return recv_dt, emaily
    
    except Exception as e:
        print(f"[MSG] Chyba při zpracování {cesta_k_souboru}: {e}")
        return None, []
    finally:
        if msg:
            msg.close()  # Přidat zavření

# --- Měření času ---
start_time = time.time()  # Record start time

# --- Zpracování a zápis do CSV ---
zaznamy = []
pocet = 0

for nazev_souboru in os.listdir(slozka):
    if pocet >= max_soubory:
        break

    if not (nazev_souboru.endswith(".eml") or nazev_souboru.endswith(".msg")):
        continue

    cesta = os.path.join(slozka, nazev_souboru)
    identifikator = ziskat_id(nazev_souboru)

    # ROZBALENÍ TUPLE PRO OBA PŘÍPADY
    if nazev_souboru.endswith(".eml"):
        recv_dt, emaily = zpracuj_eml(cesta)
    else:  # .msg
        recv_dt, emaily = zpracuj_msg(cesta)

    dt_str = recv_dt.strftime('%Y-%m-%d %H:%M:%S') if recv_dt else ''
    print(f"{identifikator} | {dt_str} -> {';'.join(emaily)}")

    zaznamy.append([identifikator, dt_str] + emaily)
    pocet += 1

# --- Zápis do CSV (vystup_all.csv) ---
max_pocet_emailu = max(len(z) - 2 for z in zaznamy)  # -1 kvůli ID
hlavicka = ['id_souboru', 'datum_prijeti'] + [f'email{i+1}' for i in range(max_pocet_emailu)]

with open("vystup_all.csv", mode='w', encoding='utf-8', newline='') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerow(hlavicka)
    writer.writerows(zaznamy)

print(f"\nVýstup uložen do: vystup_all.csv")

# --- Filtrace a zápis do CSV (vystup_bez_ct_email.csv) ---
# Zachování pozic emailů v řetězci komunikace - CT emaily nahrazeny prázdnými řetězci
zaznamy_bez_ct = [
    [identifikator, datum_prijeti] + [email if not re.search(r'@ceskatelevize\.cz$', email) else "" for email in emaily]
    for identifikator, datum_prijeti, *emaily in zaznamy  # Opraveno rozbalení
]

max_pocet_emailu_bez_ct = max(len(z) - 2 for z in zaznamy_bez_ct)  # -2 kvůli ID a datu
hlavicka_bez_ct = ['id_souboru', 'datum_prijeti'] + [f'email{i+1}' for i in range(max_pocet_emailu_bez_ct)]

with open("vystup_bez_ct_email.csv", mode='w', encoding='utf-8', newline='') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerow(hlavicka_bez_ct)
    writer.writerows(zaznamy_bez_ct)

print(f"\nVýstup uložen do: vystup_bez_ct_email.csv")

# --- Výpis délky běhu ---
end_time = time.time()  # Record end time
print(f"Délka běhu skriptu: {end_time - start_time:.2f} sekund")
print(f"Zpracováno {pocet} souborů")
print(f"Výstup uložen do: vystup_all.csv")
print(f"Výstup uložen do: vystup_bez_ct_email.csv")
print(f"Délka běhu: {end_time - start_time:.2f} s")
