"""
convert_sprint_to_excel.py  (v2 - SRM + FIT Support)

Konvertiert eine Sprint-Messdatei (.srm oder .fit) in das Excel-Format
wie im "Sprint Kohmann" (Data_Sprint Sheet), das von app.py eingelesen werden kann.

Das SRM7 Format speichert die Leistungswerte bereits in Watt.
Für die Konvertierung werden Slope und Zero-Offset benötigt, die auf dem
SRM-Gerät kalibriert wurden.

WICHTIG: Der SRM-Wert ist der bereits berechnete Leistungswert (Watt).
Die Formel lautet: P [W] = (raw - zero_offset) * slope
wobei raw die Rohfrequenz des SRM-Kraftsensors in Hz ist.

Standard-Werte (falls unbekannt):
  Slope: 16.964 Hz/Nm  (typisch für SRM PC7)
  Zero-Offset: 488 Hz  (muss vom Gerät abgelesen werden!)

Ausgabe-Format (entspricht Sprint Kohman.xlsx):
  Spalte A: Leistung [W]
  Spalte B: Herzfrequenz [bpm]  
  Spalte C: Trittfrequenz [rpm]
  Spalte D: Geschwindigkeit [km/h]
  Spalte E: Zeit als String "0:00:00.10"
  Spalte F: Zeilennummer (Index)

Zeilen 1-3: Header-Informationen
Datendaten beginnen ab Zeile 4.

Verwendung:
  python convert_sprint_to_excel.py <input.srm> [--slope 16.964] [--offset 488] [output.xlsx]
  python convert_sprint_to_excel.py <input.fit> [output.xlsx]
"""

import sys
import os
import struct
import datetime
import io
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


# ---------------------------------------------------------------------------
# SRM7 Parser
# ---------------------------------------------------------------------------

def parse_srm7_header(data):
    """Liest den SRM7-Header und gibt ein Dict mit Metadaten zurück."""
    if data[:4] != b'SRM7':
        raise ValueError("Keine gültige SRM7-Datei (falsche Magic Bytes)")

    days_raw   = struct.unpack_from('<H', data, 4)[0]
    wheel_circ = struct.unpack_from('<H', data, 6)[0]   # mm
    rec_num    = data[8]    # Aufzeichnungsintervall Zähler
    rec_den    = data[9]    # Aufzeichnungsintervall Nenner
    rec_int    = rec_num / rec_den  # Sekunden pro Sample (z.B. 0.1)
    n_blocks   = struct.unpack_from('<H', data, 10)[0]  # Anzahl 6-Byte-Datenblöcke
    n_markers  = struct.unpack_from('<H', data, 12)[0]  # Anzahl zusätzlicher Marker (Gesamtmarker = n_markers + 1)
    comment_len = data[14]                              # Länge des Notiz-Strings

    # Datum berechnen (Tage seit 01.01.1880)
    base_date = datetime.date(1880, 1, 1)
    record_date = base_date + datetime.timedelta(days=days_raw)

    # Athletenname (Header-Feld ab Offset 86, 255 Bytes im ersten Marker)
    athlete_name = data[86:150].split(b'\x00')[0].decode('latin1', errors='replace').strip()
    # Notiz / Comment (Offset 15..15+comment_len)
    if 0 < comment_len <= 70:
        notes = data[15:15 + comment_len].decode('latin1', errors='replace').strip()
    else:
        notes = data[15:84].split(b'\x00')[0].decode('latin1', errors='replace').strip()

    # Block-Tabelle und Kalibrierungs-Header auslesen
    block_base = 86 + (n_markers + 1) * 270
    calib_off  = block_base + n_blocks * 6
    blocks     = []
    zero_off_hdr = 0
    slope_hdr    = 0
    rec_cnt_hdr  = 0
    if calib_off + 6 <= len(data):
        for b_idx in range(n_blocks):
            boff = block_base + b_idx * 6
            hsec_midnight, chunk_cnt = struct.unpack_from('<IH', data, boff)
            blocks.append((hsec_midnight, chunk_cnt))
        zero_off_hdr, slope_hdr, rec_cnt_hdr = struct.unpack_from('<HHH', data, calib_off)

    return {
        'rec_int':       rec_int,
        'wheel_circ_mm': wheel_circ,
        'n_markers':     n_markers,
        'n_blocks':      n_blocks,
        'n_chunks':      rec_cnt_hdr if rec_cnt_hdr > 0 else n_blocks,
        'record_date':   record_date,
        'athlete_name':  athlete_name,
        'notes':         notes,
        'marker_offset': 86,
        'marker_size':   270,
        'block_size':    6,
        'blocks':        blocks,
        'zero_off_hdr':  zero_off_hdr,
        'slope_hdr':     round(slope_hdr / 305.57, 2) if slope_hdr > 0 else 0.0,
        'rec_cnt_hdr':   rec_cnt_hdr,
        'data_offset':   calib_off + 6,
    }


def _find_srm7_data_start(data):
    """
    Findet den exakten Beginn der 14-Byte-Datenrecords in einer SRM7-Datei.

    Strategie 1 (Exakte SRM7-Binärspezifikation):
      86 Bytes Dateiheader
      + (n_markers + 1) * 270 Bytes Marker-Tabelle
      + n_blocks * 6 Bytes Block-Zeitstempel-Tabelle
      + 6 Bytes Kalibrierung (zero_offset, slope, rec_cnt)
    """
    RECORD_SIZE = 14
    if len(data) >= 16 and data[:4] == b'SRM7':
        n_blocks  = struct.unpack_from('<H', data, 10)[0]
        n_markers = struct.unpack_from('<H', data, 12)[0]
        calib_off = 86 + (n_markers + 1) * 270 + n_blocks * 6
        data_off  = calib_off + 6
        if calib_off + 6 <= len(data):
            _, _, rec_cnt = struct.unpack_from('<HHH', data, calib_off)
            if rec_cnt > 0 and data_off + rec_cnt * RECORD_SIZE <= len(data):
                return data_off

    # Strategie 2 (Fallback-Signatursuche):
    SIGNATURES = [
        bytes([0x00, 0x82, 0x00, 0x00, 0x00, 0x8C]),
        bytes([0x00, 0x28, 0x00, 0x00, 0x00, 0x8C]),
    ]
    MIN_CONSECUTIVE = 8

    n_markers_real = struct.unpack_from('<H', data, 12)[0]
    search_from    = 154 + n_markers_real * 270
    search_to      = min(search_from + 1200, len(data) - RECORD_SIZE)

    for start in range(search_from, search_to):
        for sig in SIGNATURES:
            n_valid = 0
            for i in range(MIN_CONSECUTIVE + 5):
                off = start + i * RECORD_SIZE
                if off + RECORD_SIZE > len(data):
                    break
                if data[off + 8: off + 13] == sig[:5]:
                    n_valid += 1
                else:
                    break
            if n_valid >= MIN_CONSECUTIVE:
                return start

    n_markers_byte14 = data[14]
    n_blocks_hdr     = struct.unpack_from('<H', data, 12)[0]
    return 86 + n_markers_byte14 * 270 + n_blocks_hdr * 6


def parse_srm7_data(data, slope=1.0, zero_offset=0.0, trim_zeros=True):
    """
    Liest alle Datenpunkte aus einer SRM7-Datei (universell: Ergometer + Portable).

    14-Byte-Record-Format (SRM7 Spezifikation):
      Byte 0    : Status-Byte
      Bytes 1-2 : Leistung [W] als uint16 LE
      Byte 3    : Trittfrequenz [rpm] als uint8
      Byte 4    : Herzfrequenz [bpm] als uint8
      Bytes 5-8 : Geschwindigkeit [mm/s] als int32 LE (km/h = mm/s * 3.6 / 1000)
      Bytes 9-12: Höhe [m] als int32 LE
      Byte 13   : Temperatur [0.1 °C] als int8

    trim_zeros=True: Führende UND nachfolgende Null-Leistungswerte (Vor-/Nachlauf ohne Leistung)
    werden entfernt; die Zeit beginnt exakt bei 0:00:00.00 am ersten Leistungswert und nutzt
    die exakten SRM7-Block-Zeitstempel (identisch zum SRM-ASCII/Excel-Export).

    Returns: (records, hdr)
      records : Liste von Dicts mit elapsed_s, step_s, t_abs, power_w, cadence, heart_rate, speed_kmh
      hdr     : Dict mit Metadaten aus dem Header
    """
    hdr        = parse_srm7_header(data)
    rec_int    = hdr['rec_int']
    RECORD_SZ  = 14

    data_offset  = _find_srm7_data_start(data)
    data_section = data[data_offset:]
    n_records    = len(data_section) // RECORD_SZ
    if hdr.get('rec_cnt_hdr', 0) > 0:
        n_records = min(n_records, hdr['rec_cnt_hdr'])

    # Block-Zeitstempel für jeden Record rekonstruieren (exakt wie SRMWin Export)
    rec_abs_times = []
    for hsec_midnight, chunk_cnt in hdr.get('blocks', []):
        base_sec = hsec_midnight / 100.0
        for k in range(chunk_cnt):
            rec_abs_times.append(base_sec + k * rec_int)

    records = []
    for i in range(n_records):
        b = data_section[i * RECORD_SZ: (i + 1) * RECORD_SZ]
        if len(b) < RECORD_SZ:
            break

        raw_power   = struct.unpack_from('<H', b, 1)[0]
        raw_cadence = b[3]

        if raw_power > 0:
            power_w = (raw_power - zero_offset) * slope
            power_w = max(0.0, power_w)
        else:
            power_w = 0.0

        t_abs = rec_abs_times[i] if i < len(rec_abs_times) else (i * rec_int)

        records.append({
            'elapsed_s':   round(i * rec_int, 2),
            'step_s':      round(i * rec_int, 2),
            't_abs':       t_abs,
            'power_w':     round(power_w),
            'raw_power':   raw_power,
            'cadence':     raw_cadence,
            'raw_cadence': raw_cadence,
        })

    if trim_zeros and records:
        first_power_idx = next(
            (i for i, r in enumerate(records) if r['power_w'] > 0), None
        )
        last_power_idx = next(
            (i for i in range(len(records) - 1, -1, -1) if records[i]['power_w'] > 0), None
        )
        if first_power_idx is not None and last_power_idx is not None and last_power_idx >= first_power_idx:
            records = records[first_power_idx: last_power_idx + 1]

    if records:
        t0_abs = records[0]['t_abs']
        for j, r in enumerate(records):
            r['elapsed_s'] = round(r['t_abs'] - t0_abs, 2)
            r['step_s']    = round(j * rec_int, 2)

    return records, hdr


def parse_fit_file(filepath, trim_zeros=True):
    """
    Liest eine .fit Datei (1s-Intervalle vom Garmin/SRM-Export).
    
    trim_zeros=True: Führende Null-Leistungswerte (Vorlauf) werden entfernt.
    Gibt alle Records mit Power, HR, Cadence zurück.
    """
    try:
        import fitparse
    except ImportError:
        raise ImportError("fitparse nicht installiert. Bitte: pip install fitparse")

    fitfile = fitparse.FitFile(filepath)
    records = []
    t0 = None

    for msg in fitfile.get_messages("record"):
        vals = msg.get_values()
        ts = vals.get("timestamp")
        if ts is None:
            continue
        if t0 is None:
            t0 = ts

        elapsed = (ts - t0).total_seconds()
        power   = int(vals.get("power") or 0)
        cadence = int(vals.get("cadence") or 0)
        hr      = int(vals.get("heart_rate") or 0)
        speed   = float(vals.get("speed") or 0.0)
        speed_kmh = round(speed * 3.6, 2)

        records.append({
            "elapsed_s":  elapsed,
            "power_w":    power,
            "cadence":    cadence,
            "heart_rate": hr,
            "speed_kmh":  speed_kmh,
        })

    # Führende und nachfolgende Null-Leistungswerte abschneiden
    if trim_zeros and records:
        first_power_idx = next(
            (i for i, r in enumerate(records) if r['power_w'] > 0), None
        )
        last_power_idx = next(
            (i for i in range(len(records) - 1, -1, -1) if records[i]['power_w'] > 0), None
        )
        if first_power_idx is not None and last_power_idx is not None and last_power_idx >= first_power_idx:
            records = records[first_power_idx: last_power_idx + 1]
            t0_fit = records[0]['elapsed_s']
            for j, r in enumerate(records):
                r['elapsed_s'] = round(r['elapsed_s'] - t0_fit, 2)

    athlete_name, test_date = _get_fit_info(filepath)
    return records, athlete_name, test_date


def _get_fit_info(filepath):
    """Extrahiert Athletenname und Datum aus dem Dateinamen."""
    basename = os.path.basename(filepath).replace(".fit", "").replace(".FIT", "")
    if " - " in basename:
        athlete_name = basename.split(" - ")[0].strip()
        date_part = basename.split(" - ")[1].strip()
        try:
            parts = date_part.split("-")
            test_date = f"{parts[2]}.{parts[1]}.{parts[0]}"
        except Exception:
            test_date = date_part
    else:
        athlete_name = basename
        test_date = datetime.datetime.now().strftime("%d.%m.%Y")
    return athlete_name, test_date


# ---------------------------------------------------------------------------
# Excel Export
# ---------------------------------------------------------------------------

def seconds_to_timestr(s):
    """Konvertiert Sekunden in '0:00:00.00' Format (identisch zum SRM-Export)."""
    s_rounded = round(float(s), 2)
    m = int(s_rounded // 60)
    rem_sec = s_rounded - m * 60
    return f"0:{m:02d}:{rem_sec:05.2f}"


def create_sprint_excel(records, output_path, athlete_name="", test_date="",
                        source_filename="", rec_int_label="0.1s (SRM)"):
    """
    Erstellt Excel-Datei im SRM-Sprint-Format (identisch zu sprint knobbe.xlsx).
    
    Spalten: A=Power[W], B=HR[bpm], C=Cadence[rpm], D=Speed[km/h], E=Zeit, F=Index (1..N)
    Zeilen 1-3: Header-Info
    Ab Zeile 4: Daten
    """
    wb = openpyxl.Workbook()
    ws = wb.active

    # Sheet-Name wie Original (Initialen + Datum)
    if athlete_name and test_date:
        parts = athlete_name.split()
        initials = "".join(p[0] for p in parts if p)
        try:
            d_parts = test_date.replace(".", "-").split("-")
            date_short = d_parts[0].zfill(2) + d_parts[1].zfill(2) + d_parts[2][-2:]
            ws.title = initials + date_short
        except Exception:
            ws.title = "Data_Sprint"
    else:
        ws.title = "Data_Sprint"

    # Header-Zeilen (Rows 1-3)
    ws.cell(row=1, column=2).value = len(records)
    ws.cell(row=1, column=4).value = rec_int_label
    ws.cell(row=1, column=5).value = athlete_name
    ws.cell(row=1, column=6).value = test_date

    ws.cell(row=2, column=2).value = source_filename
    ws.cell(row=2, column=5).value = athlete_name

    ws.cell(row=3, column=2).value = len(records)
    ws.cell(row=3, column=4).value = rec_int_label
    ws.cell(row=3, column=5).value = "ERGOTEST:"

    # Datenzeilen ab Row 4
    for i, rec in enumerate(records):
        row = i + 4
        elapsed  = rec["elapsed_s"]
        time_str = seconds_to_timestr(elapsed)
        power    = rec.get("power_w", rec.get("power", 0))
        cadence  = rec.get("cadence", 0)

        ws.cell(row=row, column=1).value = int(round(power))   # A: Leistung
        ws.cell(row=row, column=3).value = int(cadence)        # C: Kadenz
        ws.cell(row=row, column=5).value = time_str            # E: Zeit
        ws.cell(row=row, column=6).value = i + 1               # F: Index (1..N)

    # Spaltenbreiten
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 8

    wb.save(output_path)
    return ws.title


def create_sprint_excel_bytes(records, athlete_name="", test_date="",
                              source_filename="", rec_int_label="0.1s (SRM)"):
    """Wie create_sprint_excel, aber gibt Bytes zurück (für Download in Streamlit)."""
    buf = io.BytesIO()
    
    wb = openpyxl.Workbook()
    ws = wb.active

    if athlete_name and test_date:
        parts = athlete_name.split()
        initials = "".join(p[0] for p in parts if p)
        try:
            d_parts = test_date.replace(".", "-").split("-")
            date_short = d_parts[0].zfill(2) + d_parts[1].zfill(2) + d_parts[2][-2:]
            ws.title = initials + date_short
        except Exception:
            ws.title = "Data_Sprint"
    else:
        ws.title = "Data_Sprint"

    ws.cell(row=1, column=2).value = len(records)
    ws.cell(row=1, column=4).value = rec_int_label
    ws.cell(row=1, column=5).value = athlete_name
    ws.cell(row=1, column=6).value = test_date
    ws.cell(row=2, column=2).value = source_filename
    ws.cell(row=2, column=5).value = athlete_name
    ws.cell(row=3, column=2).value = len(records)
    ws.cell(row=3, column=4).value = rec_int_label
    ws.cell(row=3, column=5).value = "ERGOTEST:"

    for i, rec in enumerate(records):
        row = i + 4
        elapsed  = rec["elapsed_s"]
        time_str = seconds_to_timestr(elapsed)
        power    = rec.get("power_w", rec.get("power", 0))
        cadence  = rec.get("cadence", 0)

        ws.cell(row=row, column=1).value = int(round(power))
        ws.cell(row=row, column=3).value = int(cadence)
        ws.cell(row=row, column=5).value = time_str
        ws.cell(row=row, column=6).value = i + 1

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 8

    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# CLI Entry Point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Konvertiert .srm oder .fit Sprint-Dateien ins Excel Data_Sprint Format"
    )
    parser.add_argument("input", help="Eingabedatei (.srm oder .fit)")
    parser.add_argument("output", nargs="?", help="Ausgabe-Excel (.xlsx). Standard: <input>_Data_Sprint.xlsx")
    parser.add_argument("--slope", type=float, default=1.0,
                        help="SRM Kalibrierungs-Slope [W/count] (Standard: 1.0 = direkte Wattwerte)")
    parser.add_argument("--offset", type=float, default=0.0,
                        help="SRM Zero-Offset [counts] (Standard: 0)")
    args = parser.parse_args()

    input_path = args.input
    output_path = args.output or (os.path.splitext(input_path)[0] + "_Data_Sprint.xlsx")
    ext = os.path.splitext(input_path)[1].lower()

    print(f"Eingabedatei: {input_path}")
    print(f"Format:       {ext.upper()}")
    print()

    if ext == ".srm":
        with open(input_path, "rb") as f:
            data = f.read()
        
        records, hdr = parse_srm7_data(data, slope=args.slope, zero_offset=args.offset)
        athlete_name = hdr["athlete_name"]
        test_date = hdr["record_date"].strftime("%d.%m.%Y")
        rec_int_label = f"{hdr['rec_int']}s (SRM)"
        print(f"SRM-Version:  SRM7")
        print(f"Aufzeichnung: {hdr['rec_int']}s Intervall")
        print(f"Slope:        {args.slope}")
        print(f"Zero-Offset:  {args.offset}")

    elif ext == ".fit":
        records, athlete_name, test_date = parse_fit_file(input_path)
        rec_int_label = "1s (FIT)"
    else:
        print(f"FEHLER: Unbekanntes Format '{ext}'. Nur .srm und .fit werden unterstützt.")
        sys.exit(1)

    if not records:
        print("FEHLER: Keine Datenpunkte gefunden!")
        sys.exit(1)

    powers = [r.get("power_w", r.get("power", 0)) for r in records]
    print(f"Athlet:       {athlete_name}")
    print(f"Testdatum:    {test_date}")
    print(f"Messpunkte:   {len(records)}")
    print(f"Dauer:        {records[-1]['elapsed_s']:.2f}s")
    print(f"Max Power:    {max(powers):.0f}W")
    print()

    sheet_title = create_sprint_excel(
        records=records,
        output_path=output_path,
        athlete_name=athlete_name,
        test_date=test_date,
        source_filename=os.path.basename(input_path),
        rec_int_label=rec_int_label,
    )

    print(f"Excel gespeichert: {output_path}")
    print(f"  Sheet: '{sheet_title}'  ({len(records)} Datenzeilen ab Row 4)")
    print()
    print("Fertig! Datei kann in der App unter 'Sprint-Rohdaten Excel hochladen' verwendet werden.")


if __name__ == "__main__":
    main()
