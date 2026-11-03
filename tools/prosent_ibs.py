#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Buat rekap persentase panjang stock IBS dan kirim notifikasi ringkas ntfy."""

import argparse
import csv
import hashlib
import json
import os
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.request import Request, urlopen


SOURCE = Path("loglist_ibs_prosent.csv")
OUTPUT = Path("prosent_ibs.csv")
STATE = Path(".ntfy_state_prosent_ibs.json")

DETAIL_URL = "https://duniandis.github.io/prosent_ibs/"
NTFY_TITLE = "PROSENTASE STOCK IBS"
NTFY_MESSAGE = "Prosentase stock IBS berdasarkan panjang kayu telah diperbarui."


def decimal_value(value, *, allow_blank=False):
    text = str(value or "").strip().replace(",", ".")

    if not text and allow_blank:
        return Decimal("0")

    try:
        number = Decimal(text)
    except InvalidOperation as error:
        raise ValueError(f"Nilai angka tidak valid: {value!r}") from error

    if not number.is_finite() or number < 0:
        raise ValueError(f"Nilai angka tidak valid: {value!r}")

    return number


def percentage(value, total):
    if not total:
        return ""
    return f"{Decimal(value) * Decimal('100') / Decimal(total):.2f}"


def length_label(value):
    return f"{value.normalize():f}M"


def position_sort_key(value):
    return (value.upper() != "BLOK", value.upper())


def make_summary():
    if not SOURCE.is_file():
        raise FileNotFoundError(f"File sumber tidak ditemukan: {SOURCE}")

    # posisi -> kelompok panjang -> batang, volume, volume kosong
    grouped = defaultdict(
        lambda: defaultdict(
            lambda: {
                "batang": 0,
                "volume": Decimal("0"),
                "volume_kosong": 0,
            }
        )
    )

    with SOURCE.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            raise ValueError("CSV sumber kosong atau tidak memiliki header.")

        columns = {
            (name or "").strip().lower(): name
            for name in reader.fieldnames
        }

        for required in ("nobtg", "panjang", "volume", "posisi"):
            if required not in columns:
                raise ValueError(f"Kolom CSV tidak ditemukan: {required}")

        for row_number, row in enumerate(reader, start=2):
            get = lambda name: str(row.get(columns[name]) or "").strip()
            no_btg = get("nobtg")
            posisi = get("posisi")

            if not no_btg or not posisi:
                raise ValueError(
                    f"Baris {row_number}: noBtg atau posisi kosong."
                )

            try:
                panjang = decimal_value(get("panjang"))
                volume_text = get("volume")
                volume = decimal_value(volume_text, allow_blank=True)
            except ValueError as error:
                raise ValueError(f"Baris {row_number}: {error}") from error

            if panjang <= 0:
                raise ValueError(f"Baris {row_number}: panjang harus lebih dari 0.")

            # 2.00-2.49 menjadi 2M; 2.50-2.99 menjadi 2.5M; dst.
            kelompok = (panjang * 2 // 1) / 2
            item = grouped[posisi][kelompok]
            item["batang"] += 1
            item["volume"] += volume
            item["volume_kosong"] += int(not volume_text)

    if not grouped:
        raise ValueError("Tidak ada data stock yang dapat direkap.")

    records = []

    for posisi in sorted(grouped, key=position_sort_key):
        per_panjang = grouped[posisi]
        total_batang = sum(item["batang"] for item in per_panjang.values())
        total_volume = sum(
            (item["volume"] for item in per_panjang.values()),
            Decimal("0"),
        )

        for panjang in sorted(per_panjang):
            item = per_panjang[panjang]
            records.append(
                {
                    "posisi": posisi,
                    "panjang_m": f"{panjang.normalize():f}",
                    "kelompok_panjang": length_label(panjang),
                    "batang": item["batang"],
                    "volume_m3": f'{item["volume"]:.4f}',
                    "persen_batang": percentage(item["batang"], total_batang),
                    "persen_volume": percentage(item["volume"], total_volume),
                    "total_batang_posisi": total_batang,
                    "total_volume_posisi_m3": f"{total_volume:.4f}",
                    "volume_kosong": item["volume_kosong"],
                }
            )

    with OUTPUT.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(records[0]))
        writer.writeheader()
        writer.writerows(records)

    print(f"Rekap selesai: {len(records)} kelompok panjang -> {OUTPUT}")


def file_hash(path):
    hasher = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def publish_notification():
    if not OUTPUT.is_file():
        raise FileNotFoundError(f"File rekap tidak ditemukan: {OUTPUT}")

    summary_hash = file_hash(OUTPUT)

    if STATE.is_file():
        previous = json.loads(STATE.read_text(encoding="utf-8"))
        if previous.get("sha256") == summary_hash:
            print("Rekap belum berubah; notifikasi ntfy dilewati.")
            return

    topic = os.environ.get("NTFY_TOPIC", "").strip()
    server = os.environ.get("NTFY_SERVER", "https://ntfy.sh").strip().rstrip("/")
    token = os.environ.get("NTFY_TOKEN", "").strip()

    if not topic:
        raise ValueError("NTFY_TOPIC kosong.")
    if not server:
        raise ValueError("NTFY_SERVER kosong.")

    # Notifikasi sengaja tidak memuat rincian posisi atau angka stock.
    payload = {
        "topic": topic,
        "title": NTFY_TITLE,
        "message": NTFY_MESSAGE,
        "click": DETAIL_URL,
        "actions": [
            {
                "action": "view",
                "label": "Lihat Detail",
                "url": DETAIL_URL,
            }
        ],
    }

    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    request = Request(
        server,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )

    with urlopen(request, timeout=60) as response:
        result = json.load(response)

    if result.get("event") != "message" or not result.get("id"):
        raise RuntimeError("Server ntfy belum mengonfirmasi pengiriman pesan.")

    # Status hanya disimpan setelah notifikasi benar-benar berhasil dikirim.
    STATE.write_text(
        json.dumps({"sha256": summary_hash}, indent=2) + "\n",
        encoding="utf-8",
    )
    print("Notifikasi ntfy ringkas berhasil dikirim.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--notify", action="store_true")
    args = parser.parse_args()

    make_summary()
    if args.notify:
        publish_notification()


if __name__ == "__main__":
    main()
