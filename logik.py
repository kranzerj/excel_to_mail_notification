import csv
import datetime
from collections import defaultdict
import os
import subprocess

#Josy mit ChatGPT 2025-01   Das Logging funktioniert nicht gscheit. SMTP Debug also am besten mit Serverlog oder Sniffing...


# Konfiguration
email_config = {
    "smtp_server": "hier Mailserver",
    "port": 25,
    "use_tls": False,
    "use_auth": False,
    "username": "hierBenutzername",
    "password": "hierPasswort",
    "sender_email": hier Absender Adresse",
    "enable_email": True
}
email_suffix = "@countit.at"
#base_path = r"C:\TS Benachrichtiger"
#blat_path = os.path.join(base_path, "blat_current", "blat.exe")
blat_path = os.path.join("blat_current", "blat.exe")

# Datei und aktuelles Datum
datei_name = "System und Update Liste.csv"
uebersicht_datei = "\u00dcbersicht Handlungsbedarf.csv"
aktuelles_datum = datetime.datetime.now()

# Funktion zur Verarbeitung der CSV-Datei
def verarbeite_csv(datei_name):
    ueberfaellig = defaultdict(list)
    faellig_dieser_monat = defaultdict(list)
    ohne_datum = defaultdict(list)
    ohne_ansprechpartner = defaultdict(list)

    with open(datei_name, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)  # Überspringen der Kopfzeile

        for zeile in reader:
            try:
                # Zeilen filtern
                if not zeile[0] or not zeile[1] or "x" in zeile[15]:
                    continue

                # Daten aus relevanten Spalten extrahieren
                kunde = zeile[0].strip()
                systemname = zeile[1].strip()
                typ = zeile[2].strip()
                verantwortlicher = zeile[9].strip()

                if not verantwortlicher:
                    ohne_ansprechpartner["Keine Ansprechpartner"].append((kunde, systemname, typ))
                    continue

                try:
                    monate_bis_update = int(zeile[3].strip())
                    datum_str = zeile[5].strip()
                    monat, jahr = map(int, datum_str.split('.'))
                    letztes_update = datetime.datetime(year=jahr, month=monat, day=1)

                    # Berechnung des nächsten Updates
                    neuer_monat = monat + monate_bis_update
                    neues_jahr = jahr + (neuer_monat - 1) // 12
                    neuer_monat = (neuer_monat - 1) % 12 + 1
                    naechstes_update = datetime.datetime(year=neues_jahr, month=neuer_monat, day=1)

                    # Überfällige und fällige Updates filtern
                    if naechstes_update < aktuelles_datum:
                        ueberfaellig[verantwortlicher].append((kunde, systemname, typ, naechstes_update.strftime("%m.%Y")))
                    elif naechstes_update.year == aktuelles_datum.year and naechstes_update.month == aktuelles_datum.month:
                        faellig_dieser_monat[verantwortlicher].append((kunde, systemname, typ, naechstes_update.strftime("%m.%Y")))
                except Exception:
                    # Zeilen ohne korrektes Datum
                    ohne_datum[verantwortlicher].append((kunde, systemname, typ))
            except Exception as e:
                # Fehlerhafte Zeilen ignorieren
                continue

    return ueberfaellig, faellig_dieser_monat, ohne_datum, ohne_ansprechpartner

# Ergebnisse formatieren und in Datei schreiben
def schreibe_uebersicht(ueberfaellig, faellig_dieser_monat, ohne_datum, ohne_ansprechpartner):
    with open(uebersicht_datei, "w", newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Verantwortlicher", "Kunde", "Systemname", "Typ", "Grund"])

        for verantwortlicher, systeme in ueberfaellig.items():
            for kunde, systemname, typ, naechstes_update in systeme:
                writer.writerow([verantwortlicher, kunde, systemname, typ, "Überfällig: " + naechstes_update])

        for verantwortlicher, systeme in faellig_dieser_monat.items():
            for kunde, systemname, typ, naechstes_update in systeme:
                writer.writerow([verantwortlicher, kunde, systemname, typ, "Fällig diesen Monat: " + naechstes_update])

        for verantwortlicher, systeme in ohne_datum.items():
            for kunde, systemname, typ in systeme:
                writer.writerow([verantwortlicher, kunde, systemname, typ, "Datum nicht berechenbar"])

        for kategorie, systeme in ohne_ansprechpartner.items():
            for kunde, systemname, typ in systeme:
                writer.writerow([kategorie, kunde, systemname, typ, "Kein Ansprechpartner"])

# Mailbenachrichtigung erstellen und senden
def sende_emails(ueberfaellig, faellig_dieser_monat, ohne_datum, ohne_ansprechpartner):
    if not email_config["enable_email"]:
        return

    for verantwortlicher in set(ueberfaellig.keys()) | set(faellig_dieser_monat.keys()) | set(ohne_datum.keys()):
        email = verantwortlicher + email_suffix
        inhalt = (
            "Servus,\n\n"
            "laut Update Excel hast du bei folgenden Systemen was zu tun. Bitte schau, dass das nicht immer mehr wird.\n"
            "Anbei ist auch die Gesamtübersicht. Eventuell wirfst du auch dort mal einen Blick rein.\n\n"
        )

        if verantwortlicher in ueberfaellig:
            inhalt += "\nÜberfällige Systeme:\n"
            for kunde, systemname, typ, naechstes_update in ueberfaellig[verantwortlicher]:
                inhalt += f"  - Kunde: {kunde}, System: {systemname}, Typ: {typ}, Nächstes Update: {naechstes_update}\n"

        if verantwortlicher in faellig_dieser_monat:
            inhalt += "\nFällige Systeme diesen Monat:\n"
            for kunde, systemname, typ, naechstes_update in faellig_dieser_monat[verantwortlicher]:
                inhalt += f"  - Kunde: {kunde}, System: {systemname}, Typ: {typ}, Nächstes Update: {naechstes_update}\n"

        if verantwortlicher in ohne_datum:
            inhalt += "\nSysteme ohne gültiges Datum:\n"
            for kunde, systemname, typ in ohne_datum[verantwortlicher]:
                inhalt += f"  - Kunde: {kunde}, System: {systemname}, Typ: {typ}\n"

        # E-Mail senden
        cmd = [
            blat_path,
            "-to", email,
            "-subject", "folgende Hacken hast du laut Update Excel",
            "-body", inhalt,
            "-attach", uebersicht_datei,
            "-server", email_config["smtp_server"],
            "-f", email_config["sender_email"],
            "-port", str(email_config["port"]),
        ]

        if email_config["use_tls"]:
            cmd.append("-tls")
        if email_config["use_auth"]:
            cmd.extend(["-u", email_config["username"], "-pw", email_config["password"]])

        subprocess.run(cmd)

# Hauptprogramm
if __name__ == "__main__":
    ueberfaellig, faellig_dieser_monat, ohne_datum, ohne_ansprechpartner = verarbeite_csv(datei_name)
    schreibe_uebersicht(ueberfaellig, faellig_dieser_monat, ohne_datum, ohne_ansprechpartner)
    sende_emails(ueberfaellig, faellig_dieser_monat, ohne_datum, ohne_ansprechpartner)
