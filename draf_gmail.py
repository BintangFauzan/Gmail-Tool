"""
buat_draf_gmail.py
Membuat draf Gmail (dengan lampiran PDF asli) untuk setiap lamaran di folder 'hasil'.
Email TIDAK dikirim. Anda memeriksa dan menekan Kirim sendiri di Gmail.

Jalankan SETELAH buat_lamaran.py (yang membuat hasil/draf_email.json).

Pemakaian:
    python buat_draf_gmail.py            # buat draf yang belum pernah dibuat
    python buat_draf_gmail.py --ulang    # buat ulang semua, termasuk yang sudah pernah
"""

import argparse
import base64
import json
import sys
from email.message import EmailMessage
from pathlib import Path

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ============================================================
# PENGATURAN
# ============================================================
FOLDER = Path(__file__).resolve().parent
CREDENTIAL_DIR = FOLDER / "credential"      # folder berisi credentials.json
HASIL = FOLDER / "hasil"
DAFTAR_DRAF = HASIL / "draf_email.json"     # dibuat oleh buat_lamaran.py
CATATAN_TERBUAT = HASIL / "draf_terbuat.json"   # pencegah draf ganda

SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]
BATAS_MB = 18   # batas Gmail 25 MB sesudah encoding; file asli sebaiknya di bawah ini


# ============================================================
# LOGIN KE GMAIL
# ============================================================
def cari_credentials() -> Path:
    utama = CREDENTIAL_DIR / "credentials.json"
    if utama.exists():
        return utama
    kandidat = [p for p in CREDENTIAL_DIR.glob("*.json") if p.name != "token.json"]
    if len(kandidat) == 1:
        return kandidat[0]
    sys.exit(f"File credentials.json tidak ditemukan di folder: {CREDENTIAL_DIR}\n"
             "Pastikan file hasil unduhan dari Google Cloud ada di sana.")


def login_flow(token_file: Path) -> Credentials:
    flow = InstalledAppFlow.from_client_secrets_file(str(cari_credentials()), SCOPES)
    creds = flow.run_local_server(port=0)   # membuka browser untuk login
    token_file.write_text(creds.to_json(), encoding="utf-8")
    return creds


def dapatkan_service():
    token_file = CREDENTIAL_DIR / "token.json"
    creds = None
    if token_file.exists():
        creds = Credentials.from_authorized_user_file(str(token_file), SCOPES)

    if creds and creds.valid:
        pass
    elif creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            token_file.write_text(creds.to_json(), encoding="utf-8")
        except RefreshError:
            print("Sesi login habis, silakan login ulang di browser...")
            creds = login_flow(token_file)
    else:
        creds = login_flow(token_file)

    return build("gmail", "v1", credentials=creds)


# ============================================================
# MEMBUAT DRAF
# ============================================================
def susun_pesan(item: dict, pdf: Path) -> dict:
    msg = EmailMessage()
    if item.get("kepada"):
        msg["To"] = item["kepada"]
    msg["Subject"] = item["subjek"]
    msg.set_content(item["isi"])
    msg.add_attachment(
        pdf.read_bytes(), maintype="application", subtype="pdf", filename=pdf.name
    )
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"message": {"raw": raw}}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ulang", action="store_true",
                        help="buat ulang semua draf, termasuk yang sudah pernah dibuat")
    args = parser.parse_args()

    if not DAFTAR_DRAF.exists():
        sys.exit("hasil/draf_email.json belum ada. Jalankan buat_lamaran.py dulu.")

    daftar = json.loads(DAFTAR_DRAF.read_text(encoding="utf-8"))
    sudah = set()
    if CATATAN_TERBUAT.exists() and not args.ulang:
        sudah = set(json.loads(CATATAN_TERBUAT.read_text(encoding="utf-8")))

    service = dapatkan_service()
    print(f"Membuat draf untuk {len(daftar)} lamaran...\n")

    terbuat, dilewati, gagal = [], [], []
    for item in daftar:
        nama = f"{item['perusahaan']} - {item['posisi']}"
        pdf = HASIL / item["lampiran"]

        if item["lampiran"] in sudah:
            print(f"LEWATI (sudah pernah dibuat)  {nama}")
            dilewati.append(nama)
            continue
        if not pdf.exists():
            print(f"GAGAL  {nama}: file {pdf.name} tidak ada di folder hasil")
            gagal.append(nama)
            continue

        ukuran_mb = pdf.stat().st_size / (1024 * 1024)
        if ukuran_mb > BATAS_MB:
            print(f"GAGAL  {nama}: ukuran PDF {ukuran_mb:.1f} MB melebihi {BATAS_MB} MB. "
                  "Perkecil dulu berkas lengkapnya.")
            gagal.append(nama)
            continue

        try:
            service.users().drafts().create(userId="me", body=susun_pesan(item, pdf)).execute()
        except HttpError as e:
            print(f"GAGAL  {nama}: {e}")
            gagal.append(nama)
            continue

        peringatan = "" if item.get("kepada") else "  (alamat penerima kosong, isi manual di Gmail)"
        print(f"OK     {nama}{peringatan}")
        sudah.add(item["lampiran"])
        terbuat.append(nama)
        CATATAN_TERBUAT.write_text(json.dumps(sorted(sudah), ensure_ascii=False, indent=2),
                                   encoding="utf-8")

    print(f"\nSelesai. Draf dibuat: {len(terbuat)}, dilewati: {len(dilewati)}, gagal: {len(gagal)}")
    if terbuat:
        print("Buka Gmail > Draf untuk memeriksa dan mengirim.")


if __name__ == "__main__":
    main()
