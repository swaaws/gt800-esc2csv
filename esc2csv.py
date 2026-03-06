#!/usr/bin/env python3
"""
Amprobe GT-800 ESC → CSV Converter
===================================
Parst .ESC-Dateien und exportiert Messdaten als CSV.
Duplikate (mehrfach eingelesene Messungen) werden erkannt und markiert.

Unterstützte Messarten:
  - Kaltgeräteprüfung (Feld 11: PE, Iso, etc.)
  - RPE / Einzelmessungen (Feld 2: Isolationswiderstand)
  - Auto-Messung / PRCD/RCD (Feld 33/52/5/6: Ableitstrom, Leistung)

Usage:
  python3 esc2csv.py [ORDNER_ODER_DATEI] [-o output.csv] [--dupes-only] [--no-dupes]

Beispiele:
  python3 esc2csv.py .                     # Alle .ESC im aktuellen Ordner
  python3 esc2csv.py 26030609.ESC          # Einzelne Datei
  python3 esc2csv.py . -o messungen.csv    # Output-Datei angeben
  python3 esc2csv.py . --no-dupes          # Duplikate rausfiltern
"""

import csv
import hashlib
import sys
import argparse
from dataclasses import dataclass
from pathlib import Path


# Felder die Stammdaten eines Messblocks sind
STAMM_FIELDS = {
    "12": "geraet",         # GER - Gerätenummer
    "13": "ort",            # ORT
    "14": "abteilung",      # ABT
    "15": "bezeichnung",    # BEZ
    "16": "pruefart",       # Prüfart (1=Erstprüfung?)
    "17": "pruefer",        # Prüfer-Kürzel
    "18": "datum",          # Prüfdatum
    "32": "kunde",          # KD. - Kundennummer
}

# Felder die Ergebnis-Daten liefern (verschiedene Messarten)
ERGEBNIS_FIELDS = {
    "1":  "ergebnis_rpe",       # RPE Schutzleiterwiderstand (OHM, Ampere)
    "2":  "ergebnis_iso",       # RISO Isolationswiderstand (MOHM, Volt)
    "3":  "ergebnis_ipe",       # IPE Schutzleiter-Berührungsstrom (mA)
    "5":  "ergebnis_ableit",    # IB/Ableitstrom (mA, POL_POS / POL_NEG)
    "6":  "ergebnis_leistung",  # PIL Leistungsmessung (VA/KVA/KW)
    "7":  "ergebnis_diffstrom", # Differenzstrom (A, RANGE_HIGH/LOW)
    "9":  "ergebnis_selv",      # SELV/PELV Schutzkleinspannung (V)
    "11": "ergebnis_pe_iso",    # Kaltgerät komplett (PE + Iso kombiniert)
    "33": "pruefparameter",     # Prüfparameter (z.B. "243 243 PCII")
    "52": "prueftyp_detail",    # Prüftyp-Detail (z.B. "= P")
}

# Header-Felder (einmal pro Datei)
HEADER_FIELDS = {"0", "19", "20", "21", "22"}

# Felder die einen Messblock abschließen (Ergebnis vorhanden → fertig)
BLOCK_END_FIELDS = {"11", "2", "31"}

# Einzelmessungen: Gerätename verrät die Messart
GERAET_MESSART = {
    "auto":       "Auto",
    "stromzange": "Stromzange",
    "rpe":        "RPE",
    "riso":       "RISO",
    "ipe":        "IPE",
    "ib":         "IB",
    "selv":       "SELV",
    "pil":        "PIL",
}


@dataclass
class Messung:
    datei: str = ""
    modell: str = ""
    seriennummer: str = ""
    kunde: str = ""
    geraet: str = ""
    ort: str = ""
    abteilung: str = ""
    bezeichnung: str = ""
    pruefart: str = ""
    pruefer: str = ""
    datum: str = ""
    messart: str = ""           # Kaltgerät / RPE / Auto / PRCD
    pruefparameter: str = ""    # Feld 33
    ergebnis: str = ""          # Zusammengefasstes Ergebnis

    def fingerprint(self) -> str:
        """Eindeutiger Hash zur Duplikat-Erkennung."""
        key = "|".join([
            self.kunde.strip(),
            self.geraet.strip(),
            self.ort.strip(),
            self.abteilung.strip(),
            self.bezeichnung.strip(),
            self.datum.strip(),
            self.messart.strip(),
            self.ergebnis.strip(),
        ])
        return hashlib.sha256(key.encode()).hexdigest()[:12]

    def csv_row(self) -> dict:
        return {
            "Datei": self.datei,
            "Modell": self.modell,
            "Seriennr_Prüfgerät": self.seriennummer,
            "Kunde": self.kunde,
            "Gerät": self.geraet,
            "Ort": self.ort,
            "Abteilung": self.abteilung,
            "Bezeichnung": self.bezeichnung,
            "Prüfart": self.pruefart,
            "Prüfer": self.pruefer,
            "Datum": self.datum,
            "Messart": self.messart,
            "Prüfparameter": self.pruefparameter,
            "Ergebnis": self.ergebnis,
            "Fingerprint": self.fingerprint(),
            "Duplikat": "",
        }


CSV_COLUMNS = [
    "Datei", "Modell", "Seriennr_Prüfgerät", "Kunde", "Gerät",
    "Ort", "Abteilung", "Bezeichnung", "Prüfart", "Prüfer",
    "Datum", "Messart", "Prüfparameter", "Ergebnis",
    "Fingerprint", "Duplikat",
]


def _clean(value: str) -> str:
    """Ergebnis-String aufräumen."""
    return value.strip().strip("= ").strip()


def _detect_messart(ergebnis_parts: dict, geraet: str = "") -> str:
    """Erkennt die Messart anhand der vorhandenen Ergebnis-Felder und Gerätenamen."""
    # Einzelmessungen: Gerätename verrät die Messart
    geraet_lower = geraet.lower().split()[0] if geraet else ""
    # Sonderzeichen entfernen für Matching (z.B. "RPEKÄL00P" → "rpek")
    geraet_clean = "".join(c for c in geraet_lower if c.isalpha())
    for prefix, messart in GERAET_MESSART.items():
        if geraet_clean.startswith(prefix):
            return messart

    # PRCD/RCD (hat Feld 33 mit PCII + Ableitstrom)
    if ergebnis_parts.get("33", "").upper().find("PCI") >= 0:
        return "PRCD/RCD"
    # Auto (Ableitstrom + Leistung ohne expliziten Typ)
    if ergebnis_parts.get("5") and ergebnis_parts.get("6"):
        return "Auto"
    # Kaltgerät (Feld 11: PE + Iso kombiniert)
    if ergebnis_parts.get("11"):
        return "Kaltgerät"
    # Stromzange (Differenzstrom, Feld 7)
    if ergebnis_parts.get("7"):
        return "Stromzange"
    # SELV
    if ergebnis_parts.get("9"):
        return "SELV"
    # Einzelmessungen nach Ergebnis-Feld
    if ergebnis_parts.get("1"):
        return "RPE"
    if ergebnis_parts.get("2"):
        return "RISO"
    if ergebnis_parts.get("3"):
        return "IPE"
    if ergebnis_parts.get("5"):
        return "IB"
    if ergebnis_parts.get("6"):
        return "PIL"
    return "Unbekannt"


def _build_ergebnis(ergebnis_parts: dict) -> str:
    """Baut das Ergebnis aus allen gesammelten Teilen zusammen."""
    parts = []

    # Feld 11 — Schutzleiter + Isolation (Kaltgerät)
    if ergebnis_parts.get("11"):
        parts.append(f"PE/Iso: {ergebnis_parts['11']}")

    # Feld 1 — RPE Schutzleiterwiderstand
    if ergebnis_parts.get("1"):
        parts.append(f"RPE: {ergebnis_parts['1']}")

    # Feld 2 — RISO Isolationswiderstand
    if ergebnis_parts.get("2"):
        parts.append(f"RISO: {ergebnis_parts['2']}")

    # Feld 3 — IPE Schutzleiter-Berührungsstrom
    if ergebnis_parts.get("3"):
        parts.append(f"IPE: {ergebnis_parts['3']}")

    # Feld 7 — Differenzstrom
    if ergebnis_parts.get("7"):
        parts.append(f"Diff: {ergebnis_parts['7']}")

    # Feld 9 — SELV/PELV
    if ergebnis_parts.get("9"):
        parts.append(f"SELV: {ergebnis_parts['9']}")

    # Feld 52 — Prüftyp-Detail
    if ergebnis_parts.get("52"):
        parts.append(f"Typ: {ergebnis_parts['52']}")

    # Feld 5 — Ableitstrom (kann mehrfach vorkommen: POL_POS, POL_NEG)
    for v in ergebnis_parts.get("5_list", []):
        parts.append(f"Ableit: {v}")

    # Feld 6 — Leistung
    if ergebnis_parts.get("6"):
        parts.append(f"Leistung: {ergebnis_parts['6']}")

    return " | ".join(parts) if parts else ""


def _finalize(current: "Messung | None", ergebnis_parts: dict) -> "Messung | None":
    """Schließt einen Messblock ab und setzt Messart + Ergebnis."""
    if current is None:
        return None
    current.messart = _detect_messart(ergebnis_parts, current.geraet)
    current.ergebnis = _build_ergebnis(ergebnis_parts)
    current.pruefparameter = ergebnis_parts.get("33", "")
    if current.ergebnis or current.geraet:
        return current
    return None


def parse_esc(filepath: Path) -> list["Messung"]:
    """Parst eine einzelne .ESC-Datei und gibt alle Messungen zurück."""
    messungen = []
    dateiname = filepath.name

    with open(filepath, "r", encoding="latin-1") as f:
        lines = f.readlines()

    modell = ""
    seriennummer = ""
    current: Messung | None = None
    ergebnis_parts: dict = {}  # Sammelt Ergebnis-Felder pro Block

    for raw_line in lines:
        line = raw_line.rstrip("\r\n")

        # Datei-Start oder Ende
        if line.startswith(":") or line.strip() == "99":
            continue

        # Feld-Nummer und Wert trennen
        parts = line.split(None, 1)
        if not parts:
            continue
        field_nr = parts[0]
        value = parts[1] if len(parts) > 1 else ""

        # Leere Felder (z.B. "31 " ohne Wert → Block-Ende-Signal)
        if field_nr == "31":
            # Feld 31 = Ende eines Auto/PRCD-Blocks
            m = _finalize(current, ergebnis_parts)
            if m:
                messungen.append(m)
            current = None
            ergebnis_parts = {}
            continue

        # Header-Felder
        if field_nr == "0":
            modell = value.strip()
            continue
        if field_nr == "22":
            seriennummer = value.strip()
            continue
        if field_nr in ("19", "20", "21"):
            continue

        # Neuer Messblock bei Feld 32 (Kunde)
        if field_nr == "32":
            # Vorherigen Block abschließen
            m = _finalize(current, ergebnis_parts)
            if m:
                messungen.append(m)
            current = Messung(
                datei=dateiname,
                modell=modell,
                seriennummer=seriennummer,
                kunde=value.strip(),
            )
            ergebnis_parts = {}
            continue

        # Block noch nicht gestartet → anlegen
        if current is None:
            current = Messung(
                datei=dateiname,
                modell=modell,
                seriennummer=seriennummer,
            )
            ergebnis_parts = {}

        # Stammdaten-Felder
        if field_nr in STAMM_FIELDS:
            setattr(current, STAMM_FIELDS[field_nr], value.strip())
            continue

        # Ergebnis-Felder sammeln
        if field_nr in ERGEBNIS_FIELDS:
            cleaned = _clean(value)

            if field_nr == "11":
                # Kaltgerät-Ergebnis → wenn schon eins da, neuer Block
                if ergebnis_parts.get("11"):
                    m = _finalize(current, ergebnis_parts)
                    if m:
                        messungen.append(m)
                    prev = current
                    current = Messung(
                        datei=dateiname,
                        modell=modell,
                        seriennummer=seriennummer,
                        kunde=prev.kunde,
                        ort=prev.ort,
                        abteilung=prev.abteilung,
                        bezeichnung=prev.bezeichnung,
                        pruefart=prev.pruefart,
                        pruefer=prev.pruefer,
                        datum=prev.datum,
                    )
                    ergebnis_parts = {}
                ergebnis_parts["11"] = cleaned

            elif field_nr == "5":
                # Ableitstrom kann mehrfach kommen (POL_POS, POL_NEG)
                ergebnis_parts.setdefault("5_list", []).append(cleaned)
                ergebnis_parts["5"] = True

            elif field_nr in ("33", "52"):
                ergebnis_parts[field_nr] = cleaned

            elif field_nr in ("1", "2", "3", "6", "7", "9"):
                ergebnis_parts[field_nr] = cleaned

    # Letzte Messung
    m = _finalize(current, ergebnis_parts)
    if m:
        messungen.append(m)

    return messungen


def mark_duplicates(messungen: list[Messung]) -> list[dict]:
    """Markiert Duplikate anhand des Fingerprints."""
    seen: dict[str, int] = {}
    rows = []

    for m in messungen:
        row = m.csv_row()
        fp = row["Fingerprint"]
        if fp in seen:
            seen[fp] += 1
            row["Duplikat"] = f"Ja (#{seen[fp]})"
        else:
            seen[fp] = 1
            row["Duplikat"] = "Nein"
        rows.append(row)

    return rows


def main():
    parser = argparse.ArgumentParser(
        description="Amprobe GT-800 ESC → CSV Converter"
    )
    parser.add_argument(
        "input", nargs="?", default=".",
        help="ESC-Datei oder Ordner (Standard: aktueller Ordner)"
    )
    parser.add_argument(
        "-o", "--output", default="messungen.csv",
        help="Output CSV-Datei (Standard: messungen.csv)"
    )
    parser.add_argument(
        "--no-dupes", action="store_true",
        help="Duplikate komplett rausfiltern"
    )
    parser.add_argument(
        "--dupes-only", action="store_true",
        help="Nur Duplikate anzeigen"
    )
    parser.add_argument(
        "--separator", default=";",
        help="CSV-Trennzeichen (Standard: ;)"
    )

    args = parser.parse_args()
    input_path = Path(args.input)

    # Dateien sammeln
    if input_path.is_file():
        esc_files = [input_path]
    elif input_path.is_dir():
        esc_files = sorted(input_path.glob("*.ESC"))
        if not esc_files:
            esc_files = sorted(input_path.glob("*.esc"))
    else:
        print(f"❌ Pfad nicht gefunden: {input_path}", file=sys.stderr)
        sys.exit(1)

    if not esc_files:
        print(f"❌ Keine .ESC-Dateien in {input_path}", file=sys.stderr)
        sys.exit(1)

    # Alle Dateien parsen
    alle_messungen: list[Messung] = []
    for f in esc_files:
        messungen = parse_esc(f)
        alle_messungen.extend(messungen)
        print(f"  📄 {f.name}: {len(messungen)} Messungen")

    # Duplikate markieren
    rows = mark_duplicates(alle_messungen)

    # Filtern
    if args.no_dupes:
        rows = [r for r in rows if r["Duplikat"] == "Nein"]
    elif args.dupes_only:
        rows = [r for r in rows if r["Duplikat"] != "Nein"]

    # Statistik
    total = len(rows)
    dupes = sum(1 for r in rows if r["Duplikat"] != "Nein")
    unique = total - dupes

    # Messarten-Statistik
    messarten = {}
    for r in rows:
        ma = r.get("Messart", "?")
        messarten[ma] = messarten.get(ma, 0) + 1

    # CSV schreiben
    output_path = Path(args.output)
    with open(output_path, "w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(
            csvfile, fieldnames=CSV_COLUMNS,
            delimiter=args.separator, quoting=csv.QUOTE_MINIMAL
        )
        writer.writeheader()
        writer.writerows(rows)

    print(f"\n✅ {output_path} geschrieben")
    print(f"   {total} Einträge gesamt, {unique} unique, {dupes} Duplikate")
    print(f"   Messarten: {', '.join(f'{k}: {v}' for k, v in sorted(messarten.items()))}")


if __name__ == "__main__":
    main()
