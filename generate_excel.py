import json, sys
from datetime import datetime
from pathlib import Path
import pandas as pd

BASE = Path(__file__).resolve().parent
BERKAS = BASE / "Berkas"
HASIL = BASE / "hasil"

# Sumber skill: "v6" (job-multi-bot-search) atau "procedure" (job-search-procedure)
SOURCE = sys.argv[1].lower() if len(sys.argv) > 1 else "v6"
if SOURCE not in ("v6", "procedure"):
    print(f"[ERROR] Sumber tidak dikenal: {SOURCE}. Gunakan 'v6' atau 'procedure'.")
    sys.exit(1)

OUT_DIR = BERKAS / SOURCE
OUT_DIR.mkdir(parents=True, exist_ok=True)

# Input: hasil pencarian lowongan yang sudah diverifikasi (ditulis oleh skill).
INPUT = HASIL / f"lowongan_{SOURCE}.json"

KOLOM = ["perusahaan", "alamat", "posisi", "email_tujuan",
         "tanggal", "sumber_lowongan", "catatan"]


def main():
    if not INPUT.exists():
        print(f"[ERROR] Tidak ada data lowongan: {INPUT}")
        print("   Simpan hasil verifikasi ke file itu dulu (format list JSON).")
        sys.exit(1)

    with open(INPUT, encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        print(f"[ERROR] Format salah: {INPUT} harus berisi list JSON.")
        sys.exit(1)

    rows = []
    for d in data:
        if not d.get("email_tujuan"):
            continue
        rows.append({
            "perusahaan": d.get("perusahaan", ""),
            "alamat": d.get("alamat", ""),
            "posisi": d.get("posisi", ""),
            "email_tujuan": d.get("email_tujuan", ""),
            "tanggal": d.get("tanggal") or datetime.now().strftime("%Y-%m-%d"),
            "sumber_lowongan": d.get("sumber_lowongan", SOURCE),
            "catatan": d.get("catatan", ""),
        })

    if not rows:
        print("[WARN] Tidak ada lowongan dengan email_tujuan. File Excel tidak dibuat.")
        sys.exit(0)

    seen = set()
    uniq = []
    for r in rows:
        key = (r["perusahaan"].strip().lower(), r["posisi"].strip().lower())
        if key in seen:
            continue
        seen.add(key)
        uniq.append(r)

    today = datetime.now().strftime("%Y-%m-%d")
    out = OUT_DIR / f"daftar_lowongan_{today}.xlsx"
    df = pd.DataFrame(uniq, columns=KOLOM)
    df.to_excel(out, index=False)
    print(f"[OK] Dibuat: {out} ({len(uniq)} lowongan)")
    print(f"   Sumber data: {INPUT}")


if __name__ == "__main__":
    main()
