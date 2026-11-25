"""
buat_lamaran.py
Membuat surat lamaran massal: isi template Word -> ubah ke PDF (lewat Microsoft Word) -> gabung dengan berkas lengkap.

Cara pakai: lihat PETUNJUK di bawah file ini atau penjelasan di chat.
"""

import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from docx import Document
from docxtpl import DocxTemplate
from pypdf import PdfReader, PdfWriter

# ============================================================
# PENGATURAN (ubah di sini kalau nama file Anda berbeda)
# ============================================================
FOLDER = Path(__file__).resolve().parent

TEMPLATE = FOLDER / "Berkas/template_lamaran.docx"   # surat lamaran dengan tanda isian
BERKAS = FOLDER / "Berkas/CV Dan Dokumen Pendukung - Bintang Fauzan Dyan.pdf"   # CV, pas foto, ijazah, dst. (1 PDF)
DAFTAR = FOLDER / "Berkas/daftar_lowongan.xlsx"      # daftar perusahaan tujuan
HASIL = FOLDER / "hasil"                              # folder hasil

# Data diri untuk draf email
NAMA_PELAMAR = "Bintang Fauzan Dyan"
PENDIDIKAN = "S1 Teknik Informatika"
UNIVERSITAS = "Universitas Islam Negeri Sultan Syarif Kasim Riau"
TELEPON = "088271186170"
EMAIL_PELAMAR = "bintangfauzan1000@gmail.com"

BUAT_DRAF_EMAIL = True                                # False kalau tidak perlu draf email

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]


# ============================================================
# FUNGSI BANTU
# ============================================================
def tanggal_indonesia(d: date) -> str:
    return f"{d.day} {BULAN[d.month - 1]} {d.year}"


def rapikan_tanggal(nilai: str) -> str:
    """Kosong -> tanggal hari ini. Format tanggal Excel -> format Indonesia."""
    nilai = (nilai or "").strip()
    if not nilai:
        return tanggal_indonesia(date.today())
    if re.match(r"^\d{4}-\d{2}-\d{2}", nilai):
        return tanggal_indonesia(datetime.strptime(nilai[:10], "%Y-%m-%d").date())
    return nilai


def nama_aman(teks: str) -> str:
    """Ubah teks jadi nama file yang aman."""
    teks = re.sub(r"[^\w\s-]", "", teks, flags=re.UNICODE)
    return re.sub(r"\s+", "_", teks.strip())[:60] or "tanpa_nama"


def cari_libreoffice():
    for nama in ("soffice", "libreoffice"):
        jalur = shutil.which(nama)
        if jalur:
            return jalur
    return None


def konversi_semua_ke_pdf(folder: Path) -> None:
    """Ubah semua .docx di folder jadi PDF. Utama: Microsoft Word. Cadangan: LibreOffice."""
    try:
        from docx2pdf import convert
        convert(str(folder))          # Word dibuka sekali untuk semua surat
        return
    except Exception as e:
        print(f"(Konversi lewat Word gagal: {e})")

    soffice = cari_libreoffice()
    if not soffice:
        sys.exit("Tidak bisa mengubah ke PDF. Pastikan Microsoft Word terpasang dan "
                 "'pip install docx2pdf' sudah dijalankan.")
    print("Memakai LibreOffice sebagai cadangan...")
    for docx in sorted(folder.glob("*.docx")):
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(folder), str(docx)],
            check=True, capture_output=True,
        )


def gabung_pdf(surat: Path, berkas: Path, tujuan: Path) -> None:
    penulis = PdfWriter()
    for sumber in (surat, berkas):
        for halaman in PdfReader(str(sumber)).pages:
            penulis.add_page(halaman)
    with open(tujuan, "wb") as f:
        penulis.write(f)


def isi_surat_untuk_email(dokumen: Path) -> str:
    """Ambil tujuan surat dan isi surat, tanpa tanggal atau tanda tangan."""
    paragraf = [p.text.strip() for p in Document(str(dokumen)).paragraphs]
    try:
        awal = next(i for i, teks in enumerate(paragraf)
                    if teks.casefold() == "kepada")
        salam = next(i for i, teks in enumerate(paragraf[awal:], awal)
                     if teks.casefold() == "dengan hormat,")
        akhir = next(i for i, teks in enumerate(paragraf[awal + 1:], awal + 1)
                     if teks.casefold() == "hormat saya,")
    except StopIteration as e:
        raise ValueError(
            "Bagian tujuan, salam, atau tanda tangan tidak ditemukan di template surat."
        ) from e

    if salam >= akhir:
        raise ValueError("Urutan bagian tujuan, salam, dan tanda tangan tidak valid.")

    isi = paragraf[awal:akhir]
    while isi and not isi[-1]:
        isi.pop()
    if not isi:
        raise ValueError("Isi surat untuk body email kosong.")
    return "\n".join(isi)


def buat_draf_email(data: dict, nama_file: str, dokumen_surat: Path) -> dict:
    """Draf email dengan tujuan dan isi dari surat lamaran."""
    perusahaan = data["perusahaan"].strip()
    posisi = data["posisi"].strip()
    email_tujuan = data.get("email_tujuan", "").strip()

    return {
        "perusahaan": perusahaan,
        "posisi": posisi,
        "kepada": email_tujuan,
        "subjek": f"Lamaran Kerja - {posisi} - {NAMA_PELAMAR}",
        "lampiran": f"{nama_file}.pdf",
        "isi": isi_surat_untuk_email(dokumen_surat) + "\n",
    }


def format_draf_teks(d: dict) -> str:
    return (
        f"=== {d['perusahaan']} - {d['posisi']} ===\n"
        f"Kepada  : {d['kepada']}\n"
        f"Subjek  : {d['subjek']}\n"
        f"Lampiran: {d['lampiran']}\n\n"
        f"{d['isi']}"
    )


# ============================================================
# PENGATURAN (ubah di sini kalau nama file Anda berbeda)
# ============================================================
FOLDER = Path(__file__).resolve().parent

TEMPLATE = FOLDER / "Berkas/template_lamaran.docx"   # surat lamaran dengan tanda isian
BERKAS = FOLDER / "Berkas/CV Dan Dokumen Pendukung - Bintang Fauzan Dyan.pdf"   # CV, pas foto, ijazah, dst. (1 PDF)
HASIL = FOLDER / "hasil"                              # folder hasil

SOURCE = sys.argv[1].lower() if len(sys.argv) > 1 else "v6"
if SOURCE not in ("v6", "procedure"):
    print(f"[ERROR] Sumber tidak dikenal: {SOURCE}. Gunakan 'v6' atau 'procedure'.")
    sys.exit(1)

# Cari daftar lowongan terbaru (format: daftar_lowongan_YYYY-MM-DD.xlsx) di folder sumber
def cari_daftar_terbaru() -> Path:
    target_folder = FOLDER / "Berkas" / SOURCE
    pola = "daftar_lowongan_*.xlsx"
    files = sorted(target_folder.glob(pola), reverse=True)
    if files:
        return files[0]
    # Tidak ada file ber-tanggal -> keluar dengan pesan
    print(f"[ERROR] Tidak ada data Excel terbaru di folder {SOURCE}/ (format daftar_lowongan_YYYY-MM-DD.xlsx).")
    print("   Jalankan generate_excel.py <sumber> terlebih dahulu setelah skill job search selesai.")
    sys.exit(1)

DAFTAR = cari_daftar_terbaru()


def baca_daftar() -> pd.DataFrame:
    """Baca Excel sumber sebagai teks; sel kosong jadi string kosong."""
    return pd.read_excel(DAFTAR, dtype=str).fillna("")


# ============================================================
# PROGRAM UTAMA
# ============================================================
def main() -> None:
    for f in (TEMPLATE, BERKAS, DAFTAR):
        if not f.exists():
            sys.exit(f"File tidak ditemukan: {f.name}\nPastikan ada di folder yang sama dengan skrip ini.")

    HASIL.mkdir(exist_ok=True)
    daftar = baca_daftar()
    print(f"Memproses {len(daftar)} lowongan...\n")

    draf, gagal, dibuat = [], [], []

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # TAHAP 1: isi template untuk semua lowongan
        for i, baris in daftar.iterrows():
            perusahaan = baris["perusahaan"].strip()
            posisi = baris["posisi"].strip()
            try:
                data = {k: str(v).strip() for k, v in baris.items()}
                data["tanggal"] = rapikan_tanggal(data.get("tanggal", ""))
                dasar = f"Lamaran_{nama_aman(perusahaan)}_{nama_aman(posisi)} - {NAMA_PELAMAR}"

                tpl = DocxTemplate(str(TEMPLATE))
                tpl.render(data)
                tpl.save(str(tmp / f"{dasar}.docx"))
                dibuat.append((dasar, perusahaan, posisi, data))
            except Exception as e:
                print(f"[{i + 1}] GAGAL isi template {perusahaan}: {e}")
                gagal.append(perusahaan)

        # TAHAP 2: ubah semua surat ke PDF sekaligus
        print("Mengubah surat ke PDF (Word akan terbuka sebentar)...")
        konversi_semua_ke_pdf(tmp)

        # TAHAP 3: gabung tiap surat dengan berkas lengkap
        for dasar, perusahaan, posisi, data in dibuat:
            try:
                surat_pdf = tmp / f"{dasar}.pdf"
                if not surat_pdf.exists():
                    raise FileNotFoundError("PDF surat tidak terbentuk")
                gabung_pdf(surat_pdf, BERKAS, HASIL / f"{dasar}.pdf")
                print(f"OK    {dasar}.pdf")

                if BUAT_DRAF_EMAIL:
                    draf.append(buat_draf_email(data, dasar, tmp / f"{dasar}.docx"))
            except Exception as e:
                print(f"GAGAL {perusahaan}: {e}")
                gagal.append(perusahaan)

    if BUAT_DRAF_EMAIL and draf:
        (HASIL / "draf_email.txt").write_text(
            "\n\n".join(format_draf_teks(d) for d in draf), encoding="utf-8")
        # dipakai oleh buat_draf_gmail.py
        (HASIL / "draf_email.json").write_text(
            json.dumps(draf, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\nSelesai. Berhasil: {len(draf) if BUAT_DRAF_EMAIL else len(dibuat) - len(gagal)}, gagal: {len(gagal)}")
    if gagal:
        print("Yang gagal:", ", ".join(gagal))
    print(f"Hasil ada di folder: {HASIL}")


if __name__ == "__main__":
    main()
