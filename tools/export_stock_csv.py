import csv, json, hashlib, os
from datetime import datetime, date
from openpyxl import load_workbook

XLSX  = "INPUT_ANGKUTAN_STOCK_NEW.xlsx"
SHEET = "POSISI TERAKHIR"

MIN_ROW = 4
MAX_ROW = 10000

# 1-based column index
COL_JENIS  = 8   # H
COL_VOL    = 13  # M
COL_KELAS  = 18  # R
COL_TGL    = 19  # S
COL_POSISI = 20  # T
COL_NOBTG  = 2   # B

OUT_CSV = "stock.csv"
STATE   = ".sync_state_stock.json"

def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def load_state():
    if not os.path.exists(STATE):
        return {}
    with open(STATE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_state(st):
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=2)

def norm_str(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s == "=":
        return ""
    return s

def parse_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = norm_str(v)
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except:
            pass
    return None

def should_skip_posisi(posisi_raw) -> bool:
    s = norm_str(posisi_raw).upper()
    if not s:
        return True
    if s == "DKDS":
        return True
    if "MILIR" in s:
        return True
    return False

def safe_float(v):
    try:
        if v is None:
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        s = norm_str(v).replace(",", ".")
        return float(s) if s else 0.0
    except:
        return 0.0

def main():
    # skip kalau Excel belum berubah
    xhash = sha256_file(XLSX)
    st = load_state()
    if st.get("xlsx_sha256") == xhash:
        print("Excel unchanged; skip export.")
        return

    wb = load_workbook(XLSX, read_only=True, data_only=True)
    if SHEET not in wb.sheetnames:
        raise SystemExit(f"Sheet '{SHEET}' tidak ditemukan. Ada: {wb.sheetnames}")
    ws = wb[SHEET]

    # detail group: (posisi, kelas, jenis)
    agg = {}  # key -> {"btg": int, "vol": float, "last": date}
    last_global = None

    for r in range(MIN_ROW, MAX_ROW + 1):
        nobtg  = norm_str(ws.cell(r, COL_NOBTG).value)
        if not nobtg:
            continue

        jenis  = norm_str(ws.cell(r, COL_JENIS).value)
        kelas  = norm_str(ws.cell(r, COL_KELAS).value)
        posisi = norm_str(ws.cell(r, COL_POSISI).value)
        tgl    = parse_date(ws.cell(r, COL_TGL).value)
        vol    = safe_float(ws.cell(r, COL_VOL).value)

        if should_skip_posisi(posisi):
            continue

        key = (posisi, kelas, jenis)
        rec = agg.get(key)
        if rec is None:
            rec = {"btg": 0, "vol": 0.0, "last": None}
            agg[key] = rec

        rec["btg"] += 1
        rec["vol"] += vol

        if tgl:
            if rec["last"] is None or tgl > rec["last"]:
                rec["last"] = tgl
            if last_global is None or tgl > last_global:
                last_global = tgl

    # buat total per posisi + total global
    pos_tot = {}   # posisi -> {"btg": int, "vol": float, "last": date}
    glob_btg = 0
    glob_vol = 0.0

    for (posisi, _kelas, _jenis), rec in agg.items():
        pt = pos_tot.get(posisi)
        if pt is None:
            pt = {"btg": 0, "vol": 0.0, "last": None}
            pos_tot[posisi] = pt

        pt["btg"] += rec["btg"]
        pt["vol"] += rec["vol"]
        if rec["last"]:
            if pt["last"] is None or rec["last"] > pt["last"]:
                pt["last"] = rec["last"]

        glob_btg += rec["btg"]
        glob_vol += rec["vol"]

    last_g_str = last_global.strftime("%d-%m-%Y") if last_global else ""

    # tulis CSV
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["posisi","kelas_diameter","jenis","btg","volume_m3","mutasi_terakhir_posisi","mutasi_terakhir_global"])

        # urutkan detail by posisi lalu kelas lalu jenis
        items = sorted(agg.items(), key=lambda x: (x[0][0], x[0][1], x[0][2]))

        current_pos = None
        for (posisi, kelas, jenis), rec in items:
            # kalau pindah posisi, tulis TOTAL posisi sebelumnya
            if current_pos is not None and posisi != current_pos:
                pt = pos_tot[current_pos]
                pt_last = pt["last"].strftime("%d-%m-%Y") if pt["last"] else ""
                w.writerow([current_pos, "TOTAL", "", pt["btg"], round(pt["vol"], 3), pt_last, last_g_str])
                w.writerow([])  # baris kosong pemisah

            current_pos = posisi
            last_pos = rec["last"].strftime("%d-%m-%Y") if rec["last"] else ""
            w.writerow([posisi, kelas, jenis, rec["btg"], round(rec["vol"], 3), last_pos, last_g_str])

        # TOTAL posisi terakhir
        if current_pos is not None:
            pt = pos_tot[current_pos]
            pt_last = pt["last"].strftime("%d-%m-%Y") if pt["last"] else ""
            w.writerow([current_pos, "TOTAL", "", pt["btg"], round(pt["vol"], 3), pt_last, last_g_str])
            w.writerow([])

        # TOTAL GLOBAL
        w.writerow(["GLOBAL", "TOTAL", "", glob_btg, round(glob_vol, 3), last_g_str, last_g_str])

    st["xlsx_sha256"] = xhash
    save_state(st)
    print(f"Export done -> {OUT_CSV}")

if __name__ == "__main__":
    main()
