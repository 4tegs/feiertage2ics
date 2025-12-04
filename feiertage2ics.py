# ##########################################################################################
# Hans Straßgütl
#       Mein Ersatz für die langjährige REXX um FeiertagesEinträge im Kalender zu erstellen.
#       * Benötigt eine feiertage.txt als Eingabe
#       * Editiere zuerst die feiertage.txt. Eine gute Quelle ist: https://www.ferienfeiertagedeutschland.de/2026/feiertage/
#       * Fragt nach dem zu erstellenden Jahr.    
#       * Schreibt eine feiertage.ics    
#       * Drag & Drop in den emClient Kalender. So du sie dort wieder rauslöschen willst: 
#           - Kalenderdarstellung: Agenda - sortiere nach "Erstellt" ooder "Aktualisiert"
#           
# ------------------------------------------------------------------------------------------
# Version:
#	2025 12 03		Ersterstellung.
#
# ##########################################################################################


import datetime

def read_events(filename):
    events = []
    with open(filename, "r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if not line or line.startswith(";"):
                continue  # Kommentar oder leer → überspringen

            # Format: YYYY.MM.DD TEXT
            try:
                date_part, text = line.split(" ", 1)
                year, month, day = date_part.split(".")
                events.append({
                    "year": year,
                    "month": month,
                    "day": day,
                    "text": text.strip()
                })
            except ValueError:
                print(f"⚠️ Zeile übersprungen (ungültig): {line}")
                continue
    return events


def create_ics(year, events, output_file="feiertage.ics"):
    ics = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//HaSt ICS Feiertage Wandler / 001",
        "X-LOTUS-CHARSET:UTF-8",
        "METHOD:PUBLISH"
    ]

    uid_counter = 1

    for e in events:
        created = datetime.datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
        # Datum korrekt berechnen
        dt_start = datetime.date(int(year), int(e['month']), int(e['day']))
        dt_end = dt_start + datetime.timedelta(days=1)

        dtstart = dt_start.strftime("%Y%m%d")
        dtend = dt_end.strftime("%Y%m%d")

        text = e["text"]

        # Format 1: FEIERTAG: Neujahr
        if text.upper().startswith("FEIERTAG:"):
            short = text.split(":", 1)[1].strip()
            summary = f"<<Feiertag>> {short}"
            category = "Feiertage"

        # Format 2: z. B. "Frühlingsanfang"
        else:
            summary = text
            category = "Sonstiges"

        event = [
            "BEGIN:VEVENT",
            f"UID:Feiertage-{year}-{uid_counter}",
            f"CREATED:{created}",
            f"DTSTART;VALUE=DATE:{dtstart}",
            f"DTEND;VALUE=DATE:{dtend}",
            f"SUMMARY:{summary}",
            "TRANSP:TRANSPARENT",
            "CLASS:PRIVATE",
            "X-MICROSOFT-CDO-ALLDAYEVENT:TRUE",
            "X-FUNAMBOL-ALLDAY:TRUE",
            f"CATEGORIES:{category}",
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            "DESCRIPTION:Alarm",
            "TRIGGER;RELATED=START:-PT18H",
            "END:VALARM",
            "END:VEVENT"
        ]

        uid_counter += 1
        ics.extend(event)

    ics.append("END:VCALENDAR")

    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n".join(ics))

    print(f"✔ ICS-Datei erzeugt: {output_file}")


if __name__ == "__main__":
    jahr = input("Für welches Jahr sollen die Feiertage erzeugt werden? ")
    events = read_events("feiertage.txt")
    create_ics(jahr, events)
