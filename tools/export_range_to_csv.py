import csv, json, hashlib, os
from openpyxl import load_workbook

# INPUT
XLSX = "INPUT ANGKUTAN_STOCK_NEW.xlsx"
SHEET = "DATA_UKUR"

# RANGE: Y2:AH10000
MIN_COL = 25  # Y
MAX_COL = 34  # AH
MIN_ROW = 2
MAX_ROW = 10000

# OUTPUT
OUT_CSV = "loglist1.csv"

# STATE untuk deteksi perubahan
STATE = ".sync_state.json"

def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, str):
        s = v.strip()
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

def is_invalid_nobtg(x) -> bool:
    # nobtg dianggap kosong kalau: None / "" / "0" / 0 / "0.0"
    if x is None:
        return True
    if isinstance(x, (int, float)):
        return x == 0
    s = str(x).strip()
    return (s == "" or s == "0" or s == "0.0")

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

    wrote_any = False
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)

        # jauh lebih cepat daripada ws.cell() berulang
        for row in ws.iter_rows(
            min_row=MIN_ROW,
            max_row=MAX_ROW,
            min_col=MIN_COL,
            max_col=MAX_COL,
            values_only=True
        ):
            # nobtg = kolom pertama di range (Y)
            nobtg_raw = row[0]
            if is_invalid_nobtg(nobtg_raw):
                continue

            row_vals = [cell_str(v) for v in row]
            w.writerow(row_vals)
            wrote_any = True

    if not wrote_any:
        print("Warning: tidak ada baris berisi data dalam range (setelah filter nobtg).")

    st["xlsx_sha256"] = xhash
    save_state(st)
    print(f"Export done -> {OUT_CSV}")

if __name__ == "__main__":
    main()
