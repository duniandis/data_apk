import csv, json, hashlib, os
from openpyxl import load_workbook

# INPUT
XLSX = "INPUT ANGKUTAN_STOCK_NEW.xlsx"
SHEET = "POSISI TERAKHIR"

# RANGE: AH3:AQ9999
MIN_COL = 34  # AH
MAX_COL = 43  # AQ
MIN_ROW = 3
MAX_ROW = 6000

# OUTPUT (sesuai permintaan)
OUT_CSV = "loglist1.csv"

# STATE untuk deteksi perubahan (supaya tidak commit kalau Excel sama)
STATE = ".sync_state.json"

def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, str):
        s = v.strip()
        # kalau ada sel literal "=" saja, anggap kosong
        if s == "=":
            return ""
        return s
    return str(v)

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

def main():
    # skip kalau Excel tidak berubah
    xhash = sha256_file(XLSX)
    st = load_state()
    if st.get("xlsx_sha256") == xhash:
        print("Excel unchanged; skip export.")
        return

    wb = load_workbook(XLSX, read_only=True, data_only=True)
    if SHEET not in wb.sheetnames:
        raise SystemExit(f"Sheet '{SHEET}' tidak ditemukan. Ada: {wb.sheetnames}")
    ws = wb[SHEET]

    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        wrote_any = False

        for r in range(MIN_ROW, MAX_ROW + 1):
            row_vals = []
            all_empty = True
            for c in range(MIN_COL, MAX_COL + 1):
                v = ws.cell(row=r, column=c).value
                s = cell_str(v)
                row_vals.append(s)
                if s != "":
                    all_empty = False

            if all_empty:
                continue

            w.writerow(row_vals)
            wrote_any = True

        if not wrote_any:
            print("Warning: tidak ada baris berisi data dalam range.")

    st["xlsx_sha256"] = xhash
    save_state(st)
    print(f"Export done -> {OUT_CSV}")

if __name__ == "__main__":
    main()
