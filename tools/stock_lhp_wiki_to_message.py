import csv

CSV_FILE = "stock_lhp_wiki.csv"

def num(v, places=2):
    try:
        return f"{float(v):,.{places}f}".replace(",", "_").replace(".", ",").replace("_", ".")
    except:
        return str(v)

with open(CSV_FILE, newline="", encoding="utf-8-sig") as f:
    data = list(csv.DictReader(f))

stock = [r for r in data if r["tipe"] == "STOCK" and r["jenis"] != "TOTAL"]
pem = [r for r in data if r["tipe"] == "PEMILIRAN" and r["jenis"] != "TOTAL PEMILIRAN"]
stock_total = next(r for r in data if r["tipe"] == "STOCK" and r["jenis"] == "TOTAL")
pem_total = next(r for r in data if r["tipe"] == "PEMILIRAN" and r["jenis"] == "TOTAL PEMILIRAN")

lines = [
    "STOCK LHP PT WIKI",
    "",
    "STOCK LHP PER JENIS:"
]
for r in stock:
    lines.append(f"• {r['jenis']}: {int(float(r['total_btg']))} btg | {num(r['total_vol'])} m³")
lines += [
    f"TOTAL STOCK: {int(float(stock_total['total_btg']))} btg | {num(stock_total['total_vol'])} m³",
    "",
    "PEMILIRAN PER JENIS:"
]
for r in pem:
    lines.append(f"• {r['jenis']}: {int(float(r['total_btg']))} btg | {num(r['total_vol'])} m³")
lines.append(f"TOTAL PEMILIRAN: {int(float(pem_total['total_btg']))} btg | {num(pem_total['total_vol'])} m³")
print("\n".join(lines))
