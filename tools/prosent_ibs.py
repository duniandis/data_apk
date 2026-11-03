#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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
TITLE = "Prosentase Stock Kayu IBS Berdasarkan Panjang Kayu"


def angka(value):
    return Decimal(str(value).strip().replace(",", "."))


def label_panjang(value):
    return format(value.normalize(), "f") + "M"


def persentase(value, total):
    if not total:
        return ""
    return f"{Decimal(value) * 100 / Decimal(total):.2f}"


def buat_rekap():
    # posisi -> panjang kelompok -> batang, volume, volume kosong
    data = defaultdict(
        lambda: defaultdict(
            lambda: {
                "batang": 0,
                "volume": Decimal("0"),
                "volume_kosong": 0,
            }
        )
    )

    with SOURCE.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)

        columns = {
            (name or "").strip().lower(): name
            for name in (reader.fieldnames or [])
        }

        required = ["nobtg", "panjang", "volume", "posisi"]
        missing = [name for name in required if name not in columns]

        if missing:
            raise ValueError(
                "Kolom wajib tidak ditemukan: " + ", ".join(missing)
            )

        for row_number, row in enumerate(reader, start=2):
            if not any(str(v or "").strip() for v in row.values()):
                continue

            def get(name):
                return str(row.get(columns[name]) or "").strip()

            posisi = get("posisi")
            nomor = get("nobtg")

            if not posisi or not nomor:
                raise ValueError(
                    f"Baris {row_number}: noBtg atau posisi kosong."
                )

            try:
                panjang = angka(get("panjang"))
                raw_volume = get("volume")
                volume = (
                    angka(raw_volume)
                    if raw_volume
                    else Decimal("0")
                )
            except InvalidOperation as exc:
                raise ValueError(
                    f"Baris {row_number}: panjang/volume bukan angka."
                ) from exc

            if (
                not panjang.is_finite()
                or not volume.is_finite()
                or panjang <= 0
                or volume < 0
            ):
                raise ValueError(
                    f"Baris {row_number}: panjang/volume tidak valid."
                )

            # Contoh: 2.40 -> 2; 2.50 -> 2.5; 3.00 -> 3.
            kelompok = (panjang * 2 // 1) / 2

            item = data[posisi][kelompok]
            item["batang"] += 1
            item["volume"] += volume
            item["volume_kosong"] += int(not raw_volume)

    if not data:
        raise ValueError("CSV sumber tidak memiliki data.")

    posisi_urut = sorted(
        data,
        key=lambda name: (name.upper() != "BLOK", name.upper()),
    )

    records = []
    lines = [
        "Persentase berdasarkan jumlah batang per posisi.",
        "Kelompok panjang dibulatkan turun setiap 0.5 meter.",
    ]

    for posisi in posisi_urut:
        kelompok_data = data[posisi]
        total_batang = sum(x["batang"] for x in kelompok_data.values())
        total_volume = sum(
            (x["volume"] for x in kelompok_data.values()),
            Decimal("0"),
        )
        volume_kosong = sum(
            x["volume_kosong"] for x in kelompok_data.values()
        )

        lines.extend([
            "",
            posisi,
            f"Total: {total_batang} btg | {total_volume:.4f} m³",
        ])

        for panjang, item in sorted(kelompok_data.items()):
            persen_batang = persentase(item["batang"], total_batang)
            persen_volume = persentase(item["volume"], total_volume)

            records.append({
                "posisi": posisi,
                "panjang_m": format(panjang.normalize(), "f"),
                "kelompok_panjang": label_panjang(panjang),
                "batang": item["batang"],
                "volume_m3": f'{item["volume"]:.4f}',
                "persen_batang": persen_batang,
                "persen_volume": persen_volume,
                "total_batang_posisi": total_batang,
                "total_volume_posisi_m3": f"{total_volume:.4f}",
                "volume_kosong": item["volume_kosong"],
            })

            lines.append(
                f'{label_panjang(panjang)}: '
                f'{item["batang"]} btg ({persen_batang}%)'
            )

        if volume_kosong:
            lines.append(
                f"Catatan: {volume_kosong} batang belum ada volume."
            )

    with OUTPUT.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(records[0]))
        writer.writeheader()
        writer.writerows(records)

    print(f"Rekap selesai: {OUTPUT}, {len(records)} kelompok.")
    return "\n".join(lines)


def kirim_ntfy(message):
    fingerprint = hashlib.sha256(OUTPUT.read_bytes()).hexdigest()

    if STATE.exists():
        previous = json.loads(STATE.read_text(encoding="utf-8"))
        if previous.get("sha256") == fingerprint:
            print("Rekap sama dengan notifikasi terakhir; ntfy dilewati.")
            return

    topic = os.environ.get("NTFY_TOPIC", "").strip()
    if not topic or "/" in topic:
        raise ValueError(
            "Isi NTFY_TOPIC dengan nama topik saja, bukan URL."
        )

    server = os.environ.get(
        "NTFY_SERVER", "https://ntfy.sh"
    ).strip().rstrip("/") or "https://ntfy.sh"

    token = os.environ.get("NTFY_TOKEN", "").strip()

    # Jaga ukuran pesan. CSV tetap memuat seluruh kelompok.
    if len(message.encode("utf-8")) > 3500:
        shortened = (
            message.encode("utf-8")[:3300]
            .decode("utf-8", errors="ignore")
            .rsplit("\n", 1)[0]
        )
        message = shortened + "\n\nRekap lengkap: tekan Lihat Detail."

    payload = {
        "topic": topic,
        "title": TITLE,
        "message": message,
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
        raise RuntimeError("Server belum mengonfirmasi pesan ntfy.")

    # Simpan hanya setelah pengiriman berhasil.
    STATE.write_text(
        json.dumps({"sha256": fingerprint}, indent=2) + "\n",
        encoding="utf-8",
    )
    print("Notifikasi ntfy berhasil dikirim.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--notify", action="store_true")
    args = parser.parse_args()

    message = buat_rekap()

    if args.notify:
        kirim_ntfy(message)
    else:
        print(message)


if __name__ == "__main__":
    main()
