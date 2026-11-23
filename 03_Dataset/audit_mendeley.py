
import sys
import io
import csv
import zipfile
import hashlib
from pathlib import Path
from datetime import datetime, date
from collections import Counter
from openpyxl import load_workbook

def cetak(teks=""):
    print(teks)

def ringkas(nilai):
    if isinstance(nilai, (datetime, date)):
        return nilai.isoformat()
    teks = str(nilai)
    return teks if len(teks) <= 65 else teks[:62] + "..."

def audit_excel(nama, isi):
    cetak("\n" + "=" * 90)
    cetak("FILE EXCEL: " + nama)
    cetak("=" * 90)

    wb = load_workbook(io.BytesIO(isi), read_only=False, data_only=True)

    cetak("DAFTAR SHEET:")
    for ws in wb.worksheets:
        cetak(f"- {ws.title}: {ws.max_row} baris x {ws.max_column} kolom")

    for ws in wb.worksheets:
        cetak("\n" + "#" * 90)
        cetak("SHEET: " + ws.title)
        cetak("#" * 90)

        semua = list(ws.iter_rows(values_only=True))
        semua = [
            tuple(baris)
            for baris in semua
            if any(x is not None for x in baris)
        ]

        if not semua:
            cetak("Sheet kosong.")
            continue

        cetak("\nPREVIEW 12 BARIS PERTAMA:")
        for i, baris in enumerate(semua[:12], 1):
            cetak(f"Baris {i}: {[ringkas(x) if x is not None else None for x in baris]}")

        kandidat = []
        for i, baris in enumerate(semua[:30]):
            terisi = [x for x in baris if x is not None]
            teks = sum(isinstance(x, str) for x in terisi)
            if len(terisi) >= 2:
                kandidat.append((len(terisi), teks, -i, i, baris))

        if not kandidat:
            cetak("Header tidak dapat dideteksi.")
            continue

        kandidat.sort(reverse=True)
        nomor, isi_header = kandidat[0][3], kandidat[0][4]

        while isi_header and isi_header[-1] is None:
            isi_header = isi_header[:-1]

        header = []
        for i, nilai in enumerate(isi_header):
            nama_kolom = str(nilai).strip() if nilai is not None else f"Kolom_{i+1}"
            if nama_kolom in header:
                nama_kolom += f"_{i+1}"
            header.append(nama_kolom)

        cetak(f"\nPERKIRAAN BARIS HEADER: {nomor + 1}")
        cetak("JUMLAH KOLOM:", len(header))
        for i, kolom in enumerate(header, 1):
            cetak(f"{i}. {kolom}")

        data = []
        for baris in semua[nomor + 1:]:
            baris = list(baris[:len(header)])
            baris += [None] * (len(header) - len(baris))
            if any(x is not None for x in baris):
                data.append(baris)

        cetak("\nJUMLAH BARIS DATA:", len(data))

        if not data:
            cetak("Tidak ada baris data setelah header.")
            continue

        cetak("\nPROFIL SETIAP KOLOM:")

        for i, kolom in enumerate(header):
            nilai = [baris[i] for baris in data]
            terisi = [x for x in nilai if x is not None and str(x).strip() != ""]
            kosong = len(data) - len(terisi)

            angka = [
                x for x in terisi
                if isinstance(x, (int, float)) and not isinstance(x, bool)
            ]

            tanggal = [
                x for x in terisi
                if isinstance(x, (datetime, date))
            ]

            teks = [
                x for x in terisi
                if isinstance(x, str)
            ]

            cetak(f"\n[{i+1}] {kolom}")
            cetak(f"  Terisi: {len(terisi)}")
            cetak(f"  Kosong: {kosong}")
            cetak(f"  Numerik: {len(angka)}")
            cetak(f"  Tanggal: {len(tanggal)}")
            cetak(f"  Teks: {len(teks)}")

            if angka:
                cetak(f"  Minimum: {min(angka)}")
                cetak(f"  Maksimum: {max(angka)}")

            if tanggal:
                cetak(f"  Tanggal awal: {min(tanggal)}")
                cetak(f"  Tanggal akhir: {max(tanggal)}")

            if teks:
                contoh = list(dict.fromkeys(teks))[:8]
                cetak(f"  Contoh teks: {[ringkas(x) for x in contoh]}")

        cetak("\nLIMA BARIS DATA PERTAMA:")
        for baris in data[:5]:
            cetak([ringkas(x) if x is not None else None for x in baris])

        cetak("\nKANDIDAT KOLOM WBGT DAN WAKTU:")
        for kolom in header:
            if any(kata in kolom.lower() for kata in [
                "wbgt", "timestamp", "time", "date",
                "temperature", "humidity", "globe",
                "sensor", "reference", "ambient"
            ]):
                cetak("- " + kolom)

    wb.close()

def utama():
    lokasi = Path(input("Masukkan path ZIP atau XLSX: ").strip().strip('"'))

    if not lokasi.exists():
        cetak("File tidak ditemukan.")
        return

    if lokasi.is_file() and lokasi.suffix.lower() == ".zip":
        isi_zip = zipfile.ZipFile(lokasi)
        daftar = [
            x for x in isi_zip.namelist()
            if x.lower().endswith(".xlsx") and not x.startswith("__MACOSX/")
        ]

        cetak("\nSHA-256 ZIP:")
        cetak(hashlib.sha256(lokasi.read_bytes()).hexdigest())

        cetak("\nDAFTAR EXCEL:", daftar)

        for nama in daftar:
            isi = isi_zip.read(nama)
            cetak("\nSHA-256 XLSX:")
            cetak(hashlib.sha256(isi).hexdigest())
            audit_excel(nama, isi)

        isi_zip.close()

    elif lokasi.is_file() and lokasi.suffix.lower() == ".xlsx":
        isi = lokasi.read_bytes()
        cetak("SHA-256 XLSX:", hashlib.sha256(isi).hexdigest())
        audit_excel(lokasi.name, isi)

    else:
        cetak("Format file harus ZIP atau XLSX.")

if __name__ == "__main__":
    utama()
