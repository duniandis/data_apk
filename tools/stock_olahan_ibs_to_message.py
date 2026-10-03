import csv
with open('stock_olahan_ibs.csv', newline='', encoding='utf-8-sig') as f:
    summary = next(r for r in csv.DictReader(f) if r['tipe'] == 'RINGKASAN')
print('STOCK DOKUMENT KAYU OLAHAN IBS\n(UD. SUMBER MAPAN)\n'
      f"Stock Awal : {summary['stock_awal_m3']} m³\n"
      f"Pengurangan: {summary['pengurangan_m3']} m³\n"
      f"Sisa Stock : {summary['stock_akhir_m3']} m³")
