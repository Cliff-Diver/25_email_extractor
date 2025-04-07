import os
import csv
import re
import time  # Import time module
from email import policy
from email.parser import BytesParser
import extract_msg
from email.utils import parseaddr

# --- Nastavení ---
slozka = r"C:\Users\sd232665\OneDrive - Česká televize\Plocha\Aiviro\Emaily\ExportovaneSoubory"
vystup_csv = "vystup.csv"
max_soubory = 20000

# --- Pomocné funkce ---
def ziskat_id(nazev):
    return nazev.split("---")[0] if "---" in nazev else nazev

def extrahuj_email(text):
    """Vrací čistou e-mailovou adresu z 'Jméno <email>' nebo jen 'email'"""
    return parseaddr(text)[1]

def je_validni_email(email):
    """Kontroluje, zda je e-mailová adresa validní."""
    return bool(re.match(r'^[^@\s]+@[^@\s]+\.[a-zA-Z0-9]+$', email))

def zpracuj_eml(cesta_k_souboru):
    try:
        with open(cesta_k_souboru, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
            emaily = [extrahuj_email(msg['From'])]

            # Hledání dalších odesílatelů ve formátu From: ... uvnitř těla
            if msg.is_multipart():
                body = [part.get_content() for part in msg.walk() if part.get_content_type() in ('text/plain', 'text/html')]
                obsah = "\n".join(body)
            else:
                obsah = msg.get_content()

            from_radky = re.findall(r'^From:\s+(.+?@.+?)$', obsah, flags=re.MULTILINE | re.IGNORECASE)
            for nalezeny in from_radky:
                email = extrahuj_email(nalezeny)
                if email and je_validni_email(email) and email not in emaily:
                    emaily.append(email)

            return [email for email in emaily if je_validni_email(email)]
    except Exception as e:
        print(f"[EML] Chyba při zpracování {cesta_k_souboru}: {e}")
        return []

def zpracuj_msg(cesta_k_souboru):
    try:
        msg = extract_msg.Message(cesta_k_souboru)
        emaily = [extrahuj_email(msg.sender)]

        # Hledání "From:" v textu těla zprávy (přeposlané zprávy)
        obsah = msg.body
        from_radky = re.findall(r'^From:\s+(.+?@.+?)$', obsah, flags=re.MULTILINE | re.IGNORECASE)
        for nalezeny in from_radky:
            email = extrahuj_email(nalezeny)
            if email and je_validni_email(email) and email not in emaily:
                emaily.append(email)

        return [email for email in emaily if je_validni_email(email)]
    except Exception as e:
        print(f"[MSG] Chyba při zpracování {cesta_k_souboru}: {e}")
        return []

# --- Měření času ---
start_time = time.time()  # Record start time

# --- Zpracování a zápis do CSV ---
zaznamy = []
pocet = 0

for nazev_souboru in os.listdir(slozka):
    if pocet >= max_soubory:
        break

    if nazev_souboru.endswith(".eml") or nazev_souboru.endswith(".msg"):
        cesta = os.path.join(slozka, nazev_souboru)
        identifikator = ziskat_id(nazev_souboru)

        if nazev_souboru.endswith(".eml"):
            emaily = zpracuj_eml(cesta)
        else:
            emaily = zpracuj_msg(cesta)

        print(f"{identifikator} -> {';'.join(emaily)}")
        zaznamy.append([identifikator] + emaily)
        pocet += 1

# --- Zápis do CSV (vystup_all.csv) ---
max_pocet_emailu = max(len(z) - 1 for z in zaznamy)  # -1 kvůli ID
hlavicka = ['id_souboru'] + [f'email{i+1}' for i in range(max_pocet_emailu)]

with open("vystup_all.csv", mode='w', encoding='utf-8', newline='') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerow(hlavicka)
    writer.writerows(zaznamy)

print(f"\nVýstup uložen do: vystup_all.csv")

# --- Filtrace a zápis do CSV (vystup_bez_ct_email.csv) ---
zaznamy_bez_ct = [
    [identifikator] + ([email for email in emaily if not re.search(r'@ceskatelevize\.cz$', email)] or [""])
    for identifikator, *emaily in zaznamy
]

max_pocet_emailu_bez_ct = max(len(z) - 1 for z in zaznamy_bez_ct)  # -1 kvůli ID
hlavicka_bez_ct = ['id_souboru'] + [f'email{i+1}' for i in range(max_pocet_emailu_bez_ct)]

with open("vystup_bez_ct_email.csv", mode='w', encoding='utf-8', newline='') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerow(hlavicka_bez_ct)
    writer.writerows(zaznamy_bez_ct)

print(f"\nVýstup uložen do: vystup_bez_ct_email.csv")

# --- Výpis délky běhu ---
end_time = time.time()  # Record end time
print(f"Délka běhu skriptu: {end_time - start_time:.2f} sekund")
