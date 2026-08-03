"""
lines_plan.py
=============
Mereplikasi perhitungan sheet "Lines Plan", sel B434:G458
dari file Excel Lines_Plan_-_Oil_Tanker_20704_DWT_-_Manual.xlsx.

Jalankan:
    python lines_plan.py

Output:
    calc_values.tex  (ditulis di folder yang sama dengan script ini)

Setelah selesai, compile report.tex dengan urutan:
    1. xelatex  report
    2. bibtex   report
    3. makeindex report
    4. xelatex  report
    5. xelatex  report
"""

import os
import math
from datetime import datetime

# =============================================================
# 0. Helper functions
# =============================================================

def fmt(value, decimals=4):
    """Format float untuk output LaTeX, mis. 4.2867"""
    return f"{value:.{decimals}f}"

# fmt_sci tidak dipakai di tabel ini, tapi disimpan untuk keperluan lain
# def fmt_sci(value, decimals=3):
#     """Notasi ilmiah untuk LaTeX, mis. 1.234 \\times 10^{5}"""
#     if value == 0:
#         return "0"
#     exp = int(math.floor(math.log10(abs(value))))
#     mantissa = value / (10 ** exp)
#     return f"{mantissa:.{decimals}f} \\times 10^{{{exp}}}"

# =============================================================
# 1. Data input -- Luas penampang per station
# =============================================================
# Format : { nomor_station : (Luas_A1, Luas_A2) }
# Sumber : Lines Plan - Oil Tanker 20704 DWT, sheet "Lines Plan",
#          sel B435:E457

station_data = {
    -2: (0.0000, 0.0000),
    -1: (0.0000, 0.0000),
     0: (0.0000, 0.0000),
     1: (4.2867, 4.2935),
     2: (2.6553, 2.6531),
     3: (2.3475, 2.3470),
     4: (0.9877, 0.9880),
     5: (0.3218, 0.3217),
     6: (0.4044, 0.4046),
     7: (0.1988, 0.1987),
     8: (0.1988, 0.1987),
     9: (0.1988, 0.1987),
    10: (0.1988, 0.1987),
    11: (0.1988, 0.1987),
    12: (0.1988, 0.1987),
    13: (0.1988, 0.1987),
    14: (0.1988, 0.1987),
    15: (1.4231, 1.4223),
    16: (2.3770, 2.3773),
    17: (2.5714, 2.5701),
    18: (3.6467, 3.6471),
    19: (2.0589, 2.0596),
    20: (2.3251, 2.3348),
}

print(f"Jumlah station : {len(station_data)}")
print(f"Range station  : {min(station_data)} s/d {max(station_data)}")

# =============================================================
# 2. Perhitungan selisih dan % selisih
# =============================================================
# Kolom F (Selisih)  : =ABS(D-E)
# Kolom G (% Selisih): =(F/D)*100   -- hanya jika D > 0
# G458 (SUM)         : =SUM(G435:G457)

results = []
for station, (A1, A2) in station_data.items():
    selisih = abs(A1 - A2)
    pct     = (selisih / A1 * 100) if A1 > 0 else None
    results.append({
        "station" : station,
        "luas_A1" : A1,
        "luas_A2" : A2,
        "selisih" : selisih,
        "pct"     : pct,
    })

sum_pct = sum(r["pct"] for r in results if r["pct"] is not None)

# Tampilkan tabel ke konsol
print()
print(f"{'Station':>10}  {'Luas A1':>10}  {'Luas A2':>10}  {'Selisih':>12}  {'% Selisih':>10}")
print("-" * 62)
for r in results:
    pct_str = f"{r['pct']:10.4f}" if r["pct"] is not None else f"{'---':>10}"
    print(
        f"{r['station']:>10}  {r['luas_A1']:>10.4f}  {r['luas_A2']:>10.4f}"
        f"  {r['selisih']:>12.6f}  {pct_str}"
    )
print("-" * 62)
print(f"{'SUM % Selisih':>48}  {sum_pct:>10.4f}")

# =============================================================
# 3. Ekspor ke calc_values.tex
# =============================================================

script_dir  = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(script_dir, "calc_values.tex")
timestamp   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Nama makro LaTeX tidak boleh mengandung angka (0-9) sama sekali.
# Semua nomor station dikonversi ke kata bahasa Inggris.
# Contoh: station -2 -> \LPStaMinTwo, station 10 -> \LPStaTen

STATION_WORD = {
    -2: "MinTwo",
    -1: "MinOne",
     0: "Zero",
     1: "One",
     2: "Two",
     3: "Three",
     4: "Four",
     5: "Five",
     6: "Six",
     7: "Seven",
     8: "Eight",
     9: "Nine",
    10: "Ten",
    11: "Eleven",
    12: "Twelve",
    13: "Thirteen",
    14: "Fourteen",
    15: "Fifteen",
    16: "Sixteen",
    17: "Seventeen",
    18: "Eighteen",
    19: "Nineteen",
    20: "Twenty",
}

def macro_prefix(station_number):
    """Kembalikan prefix nama makro untuk satu station (tanpa angka)."""
    return "LPSta" + STATION_WORD[station_number]

# Bangun baris-baris isi file, hindari karakter non-ASCII di output .tex
tex_lines = []

tex_lines.append("% ============================================================")
tex_lines.append("% AUTO-GENERATED oleh lines_plan.py -- " + timestamp)
tex_lines.append("% Sumber Excel : Lines_Plan_-_Oil_Tanker_20704_DWT_-_Manual.xlsx")
tex_lines.append("% Sheet        : Lines Plan, sel B434:G458")
tex_lines.append("% JANGAN EDIT MANUAL -- jalankan lines_plan.py untuk regenerasi.")
tex_lines.append("% ============================================================")
tex_lines.append("")
tex_lines.append("% -- Nilai global --------------------------------------------")
tex_lines.append("\\newcommand{\\LPSumPct}"        + "{" + fmt(sum_pct, 4)          + "}")
tex_lines.append("\\newcommand{\\LPJumlahStation}" + "{" + str(len(station_data))   + "}")
tex_lines.append("")
tex_lines.append("% -- Nilai per station ---------------------------------------")

for r in results:
    p = macro_prefix(r["station"])
    pct_val = r["pct"] if r["pct"] is not None else 0.0

    tex_lines.append("\\" + "newcommand{\\" + p + "Aone}{" + fmt(r["luas_A1"], 4) + "}")
    tex_lines.append("\\" + "newcommand{\\" + p + "Atwo}{" + fmt(r["luas_A2"], 4) + "}")
    tex_lines.append("\\" + "newcommand{\\" + p + "Sel}{"  + fmt(r["selisih"], 6)  + "}")
    tex_lines.append("\\" + "newcommand{\\" + p + "Pct}{"  + fmt(pct_val,  4)      + "}")
    tex_lines.append("")

content = "\n".join(tex_lines)

with open(output_path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)

print()
print(f"Berhasil ditulis : {output_path}")
print("-" * 60)
print(content[:600], "...")
print("-" * 60)
print()
print("Verifikasi karakter non-ASCII :", end=" ")
bad_chars = [(i, c) for i, c in enumerate(content) if ord(c) > 127]
print("TIDAK ADA" if not bad_chars else f"ADA: {bad_chars[:5]}")
print("Verifikasi double backslash   :", end=" ")
bad_bs = [ln for ln in content.split("\n") if "\\\\" in ln and not ln.startswith("%")]
print("TIDAK ADA" if not bad_bs else f"ADA: {bad_bs[:3]}")
print()
print("Sekarang compile report.tex dengan XeLaTeX -> BibTeX -> MakeIndex -> XeLaTeX -> XeLaTeX")