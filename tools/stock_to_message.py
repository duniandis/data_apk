#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

CSV_PATH = Path("stock.csv")

# ---------- helpers ----------
def parse_int(x, default=0):
    try:
        s = str(x).strip()
        if s == "":
            return default
        return int(float(s))
    except Exception:
        return default

def parse_float(x, default=0.0):
    try:
        s = str(x).strip().replace(",", ".")
        if s == "":
            return default
        return float(s)
    except Exception:
        return default

def parse_date_ddmmyyyy(s):
    """
    stock.csv kamu pakai format: 31-12-2025
    kalau kosong / invalid -> None
    """
    if s is None:
        return None
    t = str(s).strip()
    if not t:
        return None
    # normalisasi separator
    t = t.replace("/", "-").replace(".", "-")
    for fmt in ("%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(t, fmt).date()
        except Exception:
            pass
    return None

def fmt_date(d):
    return d.strftime("%d-%m-%Y") if d else "-"

def fmt_btg(n):
    return f"{n} btg"

def fmt_vol(v):
    # 2 digit biar stabil
    return f"{v:,.2f}".replace(",", "") + " m³"

# ---------- main ----------
def main():
    if not CSV_PATH.exists():
        raise SystemExit(f"File tidak ditemukan: {CSV_PATH}")

    # Struktur:
    # posisi_data[posisi] = {
    #   "jenis": OrderedDict({jenis: {"btg": int, "vol": float}}),
    #   "total_btg": int,
    #   "total_vol": float,
    #   "last_date": date|None
    # }
    posisi_data = OrderedDict()

    with CSV_PATH.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise SystemExit("stock.csv kosong / tidak ada header")

        # cari nama kolom yang dipakai (toleran kalau beda kapital/spasi)
        fields = {c.strip(): c for c in reader.fieldnames}

        def getcol(name):
            # cari exact
            if name in fields:
                return fields[name]
            # cari case-insensitive
            for k, v in fields.items():
                if k.lower() == name.lower():
                    return v
            return None

        col_posisi = getcol("posisi")
        col_jenis = getcol("jenis")
        col_btg = getcol("btg")
        col_vol = getcol("volume_m3")
        col_last_pos = getcol("mutasi_terakhir_posisi")
        col_last_global = getcol("mutasi_terakhir_global")

        # minimal wajib ada ini:
        missing = [n for n, c in [
            ("posisi", col_posisi),
            ("jenis", col_jenis),
            ("btg", col_btg),
            ("volume_m3", col_vol),
        ] if c is None]
        if missing:
            raise SystemExit(f"Kolom wajib tidak ada di stock.csv: {', '.join(missing)}")

        for row in reader:
            posisi = (row.get(col_posisi) or "").strip()
            jenis = (row.get(col_jenis) or "").strip()

            if not posisi or not jenis:
                continue

            btg = parse_int(row.get(col_btg))
            vol = parse_float(row.get(col_vol))

            # tanggal mutasi per posisi (lebih prioritas), kalau kosong pakai global
            dpos = parse_date_ddmmyyyy(row.get(col_last_pos)) if col_last_pos else None
            dglob = parse_date_ddmmyyyy(row.get(col_last_global)) if col_last_global else None
            d = dpos or dglob

            if posisi not in posisi_data:
                posisi_data[posisi] = {
                    "jenis": OrderedDict(),
                    "total_btg": 0,
                    "total_vol": 0.0,
                    "last_date": None
                }

            p = posisi_data[posisi]

            if jenis not in p["jenis"]:
                p["jenis"][jenis] = {"btg": 0, "vol": 0.0}

            p["jenis"][jenis]["btg"] += btg
            p["jenis"][jenis]["vol"] += vol

            p["total_btg"] += btg
            p["total_vol"] += vol

            if d:
                if (p["last_date"] is None) or (d > p["last_date"]):
                    p["last_date"] = d

    if not posisi_data:
        print("📦 UPDATE STOCK\n\nTidak ada data di stock.csv")
        return

    # GLOBAL dari total posisi
    global_btg = 0
    global_vol = 0.0
    global_last = None
    for _, p in posisi_data.items():
        global_btg += p["total_btg"]
        global_vol += p["total_vol"]
        if p["last_date"]:
            if (global_last is None) or (p["last_date"] > global_last):
                global_last = p["last_date"]

    # Biar BLOK muncul paling atas kalau ada
    ordered_keys = list(posisi_data.keys())
    if "BLOK" in posisi_data:
        ordered_keys.remove("BLOK")
        ordered_keys = ["BLOK"] + ordered_keys

    # Output
    lines = []
    lines.append("📦 UPDATE STOCK")
    lines.append("")
    lines.append(f"Update terakhir (mutasi): {fmt_date(global_last)}")
    lines.append("")
    lines.append("STOCK GLOBAL")
    lines.append(f"Batang : {global_btg} btg")
    lines.append(f"Volume : {fmt_vol(global_vol)}")
    lines.append("")

    # Format kolom jenis agar rapi
    # tentukan lebar jenis maksimum (dibatasi biar gak kepanjangan)
    max_jenis_len = 0
    for pos in ordered_keys:
        for jenis in posisi_data[pos]["jenis"].keys():
            max_jenis_len = max(max_jenis_len, len(jenis))
    max_jenis_len = min(max_jenis_len, 18)  # biar aman layar hp

    sep = "=" * 16

    for pos in ordered_keys:
        p = posisi_data[pos]
        lines.append(sep)
        lines.append(pos)
        lines.append(f"Terakhir mutasi : {fmt_date(p['last_date'])}")
        lines.append(f"Total : {p['total_btg']} btg | {fmt_vol(p['total_vol'])}")
        lines.append("")

        # urutkan jenis by volume desc (lebih enak dilihat)
        items = list(p["jenis"].items())
        items.sort(key=lambda kv: kv[1]["vol"], reverse=True)

        for jenis, agg in items:
            j = jenis[:max_jenis_len]
            jpad = j.ljust(max_jenis_len)
            lines.append(f"  {jpad} : {agg['btg']:>5} btg | {fmt_vol(agg['vol']).rjust(12)}")

        lines.append("")

    print("\n".join(lines))

if __name__ == "__main__":
    main()
