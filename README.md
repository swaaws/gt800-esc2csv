# GT-800 ESC → CSV Converter

Konvertiert Messdaten des **GT-800 Geräteprüfers** vom proprietären ESC-Format in lesbares CSV. Erkennt Duplikate bei mehrfachem Einlesen automatisch.

> **Hinweis:** Das ESC-Protokoll wurde durch KI-gestützte Analyse (Claude / OpenClaw) der Rohdaten reverse-engineered, da keine offizielle Dokumentation verfügbar ist. Die Feldzuordnungen basieren auf empirischen Tests mit einem GT-800 (Firmware 8.38/8.08/2.02).

## Features

- ✅ Parst alle Messarten des GT-800 (Kaltgerät, Auto, RPE, RISO, IPE, IB, Stromzange, SELV, PIL, PRCD/RCD)
- ✅ Duplikat-Erkennung über SHA256-Fingerprint (dateiübergreifend)
- ✅ CSV mit Semikolon-Trennung (Excel-kompatibel, UTF-8 mit BOM)
- ✅ Keine externen Abhängigkeiten (Python 3.10+ Standardbibliothek)

## Installation

```bash
git clone https://github.com/swaaws/amprobe-esc2csv.git
cd amprobe-esc2csv
```

Benötigt nur **Python 3.10+** (keine pip-Pakete nötig).

## Verwendung

```bash
# Alle .ESC-Dateien im aktuellen Ordner konvertieren
python3 esc2csv.py .

# Einzelne Datei
python3 esc2csv.py 26030609.ESC

# Output-Datei angeben
python3 esc2csv.py . -o export.csv

# Nur unique Einträge (Duplikate rausfiltern)
python3 esc2csv.py . --no-dupes

# Nur Duplikate anzeigen
python3 esc2csv.py . --dupes-only

# Komma statt Semikolon
python3 esc2csv.py . --separator ","
```

## CSV-Spalten

| Spalte | Beschreibung |
|--------|-------------|
| Datei | Quelldatei (.ESC) |
| Modell | Gerätemodell (GT-800) |
| Seriennr_Prüfgerät | Seriennummer des Prüfgeräts |
| Kunde | KD. — Kundennummer |
| Gerät | GER — Gerätenummer / Prüflingskennung |
| Ort | ORT |
| Abteilung | ABT |
| Bezeichnung | BEZ |
| Prüfart | Prüfart-Code |
| Prüfer | Prüfer-Kürzel |
| Datum | Prüfdatum |
| Messart | Automatisch erkannte Messart |
| Prüfparameter | Parameter bei PRCD/Auto (Feld 33) |
| Ergebnis | Zusammengefasste Messwerte |
| Fingerprint | SHA256-Hash zur Duplikaterkennung |
| Duplikat | Nein / Ja (#N) |

## Duplikat-Erkennung

Der Fingerprint basiert auf: **Kunde + Gerät + Ort + Abteilung + Bezeichnung + Datum + Messart + Ergebnis**

Wenn dieselbe Messung in mehreren ESC-Dateien vorkommt (z.B. durch wiederholtes Auslesen vom GT-800), wird sie automatisch als Duplikat markiert. Das erste Vorkommen gilt als Original.

---

## ESC-Protokoll Dokumentation

### Dateiformat

ESC-Dateien sind zeilenbasierte Textdateien (Latin-1 Encoding). Jede Zeile besteht aus einer **Feldnummer** gefolgt von einem **Wert**, getrennt durch Leerzeichen.

```
:XXXXX              ← Geräte-Seriennummer (Datei-Header)
0 GT-800            ← Modellbezeichnung
19 8.38             ← Firmware Version 1
20 8.08             ← Firmware Version 2
21 2.02             ← Firmware Version 3
22 XXXXXXXX         ← Prüfgeräte-Seriennummer
...Messblöcke...
99                  ← Dateiende
```

### Header-Felder (einmal pro Datei)

| Feld | Bedeutung | Beispiel |
|------|-----------|---------|
| `:` | Geräte-ID (erste Zeile, mit Doppelpunkt) | `:XXXXX` |
| 0 | Modellbezeichnung | `GT-800` |
| 19 | Firmware Version 1 | `8.38` |
| 20 | Firmware Version 2 | `8.08` |
| 21 | Firmware Version 3 | `2.02` |
| 22 | Seriennummer Prüfgerät | `XXXXXXXX` |
| 99 | Dateiende-Marker | (kein Wert) |

### Stammdaten-Felder (pro Messblock)

Diese Felder beschreiben den Prüfling und werden am GT-800 über das Einstellmenü gesetzt:

| Feld | GT-800 Kürzel | Bedeutung | Beispiel |
|------|---------------|-----------|---------|
| 32 | KD. | Kundennummer | `2346` |
| 12 | GER | Gerätenummer / Prüflingskennung | `1235` |
| 13 | ORT | Prüfort | `3457` |
| 14 | ABT | Abteilung | `4568` |
| 15 | BEZ | Bezeichnung | `5679` |
| 16 | — | Prüfart (1 = Erstprüfung?) | `1` |
| 17 | — | Prüfer-Kürzel | `DHU/MHO` |
| 18 | — | Prüfdatum | `6.3.26` |

### Ergebnis-Felder (pro Messung)

| Feld | Messart | Beschreibung | Beispiel |
|------|---------|-------------|---------|
| 11 | Kaltgerät | PE + Isolation kombiniert | `= P = 0.06 OHM = 0.30 OHM > 20.00 MOHM = 1.00 MOHM = X OHM = OK` |
| 1 | RPE | Schutzleiterwiderstand | `= 0.00 OHM = 0.30 OHM = 0.07 OHM = 5.0 A` |
| 2 | RISO | Isolationswiderstand | `> 100 MOHM = 0.30 MOHM = PCIH = 250 V` |
| 3 | IPE | Schutzleiter-Berührungsstrom | `< 0.25 mA = 3.50 mA = PCI` |
| 5 | IB / Ableitstrom | Ableitstrom (kann mehrfach: POL_POS, POL_NEG) | `< 0.02 mA = 0.50 mA = POL_POS` |
| 6 | PIL | Leistungsmessung | `= 1.74 KVA = 1.74 KW = 7.99 A = 218 V = 1.00` |
| 7 | Stromzange | Differenzstrom | `< 0.2 A = RANGE_HIGH` |
| 9 | SELV | Schutzkleinspannung (PELV) | `< 10.0 V = 50.00 V = PELV` |

### Steuerfelder (PRCD/RCD und Auto)

| Feld | Bedeutung | Beispiel |
|------|-----------|---------|
| 33 | Prüfparameter | `243 243 PCII` |
| 52 | Prüftyp-Detail | `= P` |
| 31 | Block-Ende (PRCD/Auto) | (leer) |

### Ergebnis-Format

Die Ergebnis-Werte folgen einem gemeinsamen Schema:

```
[Operator] Messwert Einheit = Grenzwert Einheit [= Zusatzinfo]
```

- `=` Messwert (exakt)
- `<` Messwert unter Messbereich
- `>` Messwert über Messbereich
- `OK` / Fehlen von `FAIL` = Messung bestanden

**Beispiel Kaltgerät (Feld 11):**
```
= P = 0.06 OHM = 0.30 OHM > 20.00 MOHM = 1.00 MOHM = X OHM = OK
  │    │          │          │              │          │        └─ Gesamtergebnis
  │    │          │          │              │          └─ Berührungsstrom (X = nicht gemessen?)
  │    │          │          └──────────────└─ Isolation: Messwert > 20 MΩ, Grenze 1 MΩ
  │    └──────────└─ Schutzleiter: 0.06 Ω gemessen, Grenze 0.30 Ω
  └─ Prüfmodus (P = mit Prüfstrom?)
```

### Messblock-Struktur

Ein neuer Messblock beginnt immer mit **Feld 32 (Kunde)**. Je nach Messart folgen unterschiedliche Ergebnis-Felder:

**Kaltgerät:**
```
32 → 12 → 16 → 13 → 14 → 15 → 17 → 18 → 11
```

**PRCD/RCD und Auto:**
```
32 → 12 → 16 → 13 → 14 → 15 → 17 → 18 → 33 → 52 → 5 → 5 → 6 → 31
```

**Einzelmessungen (RPE, RISO, IPE, IB, SELV, PIL, Stromzange):**
```
32 → 12 → 16 → 13 → 14 → 15 → 17 → 18 → [1|2|3|5|6|7|9]
```

Bei Einzelmessungen enthält **Feld 12 (Gerät)** den Messart-Namen (z.B. `RISO`, `IPE`, `AUTO`, `STROMZANGE`).

### Erkannte Messarten

| Messart | Erkennung | Schalterstellung am GT-800 |
|---------|-----------|---------------------------|
| Kaltgerät | Feld 11 vorhanden | Kaltgeräteprüfung |
| PRCD/RCD | Feld 33 mit "PCI" + Gerät ≠ AUTO | PRCD/RCD-Test |
| Auto | Gerät = `AUTO` | Automatische Prüfsequenz |
| RPE | Gerät beginnt mit `RPE` | Schutzleiterwiderstand |
| RISO | Gerät = `RISO` | Isolationswiderstand |
| IPE | Gerät = `IPE` | Schutzleiter-Berührungsstrom |
| IB | Gerät = `IB` | Ableitstrom |
| Stromzange | Gerät = `STROMZANGE` oder Feld 7 | Differenzstrommessung |
| SELV | Gerät = `SELV` | Schutzkleinspannung |
| PIL | Gerät = `PIL` | Leistungsmessung |

## Einschränkungen

- Nur mit GT-800 getestet — andere Amprobe-Modelle könnten abweichende Feldnummern verwenden
- Encoding ist Latin-1 (ISO 8859-1), Umlaute in Gerätenamen können verstümmelt sein
- Ergebnis-Parsing ist textbasiert — keine Extraktion einzelner Messwerte in separate Spalten

## Lizenz

MIT

---

*Dieses Projekt und die Protokoll-Dokumentation wurden mit Unterstützung von KI (Claude / [OpenClaw](https://github.com/openclaw/openclaw)) durch Analyse realer Messdaten erstellt. Das ESC-Format wurde reverse-engineered — es gibt keine offizielle Spezifikation.*
