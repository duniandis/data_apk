import csv
from collections import defaultdict

CSV_FILE = "stock.csv"

def to_int(x):
    try:
        return int(float(str(x).strip()))
    except:
        return 0

def to_float(x):
    try:
        return float(str(x).replace(",", ".").strip())
    except:
        return 0.0

def main():
    rows = []
    with open(CSV_FILE, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    if not rows:
        print("STOCK kosong.")
        return

    # tanggal mutasi global terakhir
    mutasi_global = max((d.get("mutasi_terakhir_global","").strip() for d in rows), default="-")

    # total global
    total_btg = sum(to_int(d.get("btg","0")) for d in rows)
    total_vol = sum(to_float(d.get("volume_m3","0")) for d in rows)

    # group per posisi -> jenis
    posisi_map = defaultdict(lambda: defaultdict(lambda: {"btg": 0, "vol": 0.0}))
    for d in rows:
        posisi = (d.get("posisi") or "").strip()
        jenis = (d.get("jenis") or "").strip()
        if not posisi or not jenis:
            continue
        posisi_map[posisi][jenis]["btg"] += to_int(d.get("btg","0"))
        posisi_map[posisi][jenis]["vol"] += to_float(d.get("volume_m3","0"))

    lines = []
    lines.append("📦 UPDATE STOCK")
    lines.append("")
    lines.append(f"Update terakhir (mutasi): {mutasi_global}")
    lines.append("")
    lines.append("STOCK GLOBAL")
    lines.append(f"Batang : {total_btg} btg")
    lines.append(f"Volume : {total_vol:.2f} m³")
    lines.append("")

    for posisi in sorted(posisi_map.keys(), key=lambda s: s.upper()):
        lines.append("================")
        lines.append(posisi.upper())
        jenis_map = posisi_map[posisi]

        for jenis in sorted(jenis_map.keys(), key=lambda s: s.upper()):
            v = jenis_map[jenis]
            lines.append(f"  {jenis:<14}: {v['btg']:>6} btg | {v['vol']:>10.2f} m³")
        lines.append("")

    print("\n".join(lines))

if __name__ == "__main__":
    main()
