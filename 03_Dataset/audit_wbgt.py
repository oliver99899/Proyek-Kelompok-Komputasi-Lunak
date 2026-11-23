
import zipfile
import io
import math
import statistics
import datetime as dt
from pathlib import Path
from openpyxl import load_workbook

def angka(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None

def detik(x):
    if isinstance(x, dt.datetime):
        x = x.time()

    if isinstance(x, dt.time):
        return x.hour * 3600 + x.minute * 60 + x.second

    if isinstance(x, (int, float)):
        return int(x * 86400) % 86400

    if isinstance(x, str):
        try:
            jam = dt.datetime.strptime(x.strip(), "%H:%M:%S")
            return jam.hour * 3600 + jam.minute * 60 + jam.second
        except ValueError:
            return None

    return None

def kolom(header, kata):
    for i, nama in enumerate(header):
        nama = str(nama).lower().replace(" ", "")
        if all(k in nama for k in kata):
            return i
    return None

def statistik(pasang):
    er = [b - a for a, b in pasang]
    bias = statistics.mean(er)
    mae = statistics.mean(abs(x) for x in er)
    rmse = math.sqrt(statistics.mean(x * x for x in er))

    print("Jumlah pasangan:", len(er))
    print(f"Bias: {bias:.5f} °C")
    print(f"MAE: {mae:.5f} °C")
    print(f"RMSE: {rmse:.5f} °C")
    print(f"SD galat: {statistics.stdev(er) if len(er)>1 else 0:.5f} °C")
    print(f"Galat minimum: {min(er):.5f} °C")
    print(f"Galat maksimum: {max(er):.5f} °C")
    print("Underestimation:", sum(x < 0 for x in er))
    print("Overestimation:", sum(x > 0 for x in er))
    print("Sama:", sum(x == 0 for x in er))

def waktuinfo(waktu):
    waktu = [detik(x) for x in waktu]
    waktu = [x for x in waktu if x is not None]

    if len(waktu) < 2:
        print("Informasi waktu tidak cukup.")
        return

    selisih = [
        (b - a) % 86400
        for a, b in zip(waktu, waktu[1:])
    ]

    selisih = [x for x in selisih if x > 0]

    if not selisih:
        print("Interval tidak dapat dihitung.")
        return

    print("Waktu awal:", waktu[0])
    print("Waktu akhir:", waktu[-1])
    print("Durasi:", (waktu[-1] - waktu[0]) % 86400, "detik")
    print("Interval median:", statistics.median(selisih), "detik")
    print("Interval minimum:", min(selisih), "detik")
    print("Interval maksimum:", max(selisih), "detik")
    print("Interval >60 detik:", sum(x > 60 for x in selisih))

def utama():
    lokasi = Path(input("Path ZIP/XLSX: ").strip().strip('"'))

    if not lokasi.exists():
        print("File tidak ditemukan.")
        return

    if lokasi.suffix.lower() == ".zip":
        with zipfile.ZipFile(lokasi) as z:
            berkas = [
                x for x in z.namelist()
                if x.lower().endswith(".xlsx")
                and not x.startswith("__MACOSX/")
            ]

            if not berkas:
                print("Tidak ada XLSX.")
                return

            data = z.read(berkas[0])
    else:
        data = lokasi.read_bytes()

    wb = load_workbook(io.BytesIO(data), data_only=True)

    for nama in ["Final Lab Test 1", "Final Lab Test 2", "Final Lab Test 3"]:
        ws = wb[nama]
        baris = list(ws.iter_rows(values_only=True))
        header = [str(x).strip() if x is not None else "" for x in baris[0]]

        iw = kolom(header, ["wbgt", "reference"])
        idv = kolom(header, ["wbgt", "developedsystem"])
        it = kolom(header, ["time"])

        pasang = []
        waktu = []
        ref = []
        dev = []

        for row in baris[1:]:
            if iw is None or idv is None:
                continue

            a = angka(row[iw])
            b = angka(row[idv])

            if a is not None and b is not None:
                pasang.append((a, b))
                ref.append(a)
                dev.append(b)

                if it is not None:
                    waktu.append(row[it])

        print("\n" + "=" * 65)
        print(nama)
        print("=" * 65)

        if not pasang:
            print("Pasangan WBGT tidak ditemukan.")
            continue

        statistik(pasang)
        print("\nRENTANG PENGUKURAN")
        print(f"WBGT referensi: {min(ref):.3f}–{max(ref):.3f}")
        print(f"WBGT prototipe: {min(dev):.3f}–{max(dev):.3f}")

        print("\nJUMLAH REFERENSI DI ATAS AMBANG")
        for batas in [27.5, 28, 29, 30, 31, 32]:
            jumlah = sum(x >= batas for x in ref)
            print(f">= {batas}: {jumlah} ({jumlah/len(ref)*100:.1f}%)")

        print("\nJUMLAH REFERENSI DI SEKITAR AMBANG")
        for batas in [27.5, 28, 29, 30, 31, 32]:
            jumlah = sum(abs(x-batas) <= 0.5 for x in ref)
            print(f"{batas} ± 0.5 °C: {jumlah}")

        print("\nINFORMASI WAKTU")
        waktuinfo(waktu)

    print("\n" + "=" * 65)
    print("PEMERIKSAAN DATA LAPANGAN")
    print("=" * 65)

    for nama, wbgtnama, mode in [
        ("Final Field Test outdoor", "WBGT outdoor", "outdoor"),
        ("Final Field Test indoor", "WBGT Indoor", "indoor")
    ]:
        ws = wb[nama]
        baris = list(ws.iter_rows(values_only=True))
        header = [str(x).strip() if x is not None else "" for x in baris[0]]

        iw = kolom(header, ["wbgt"])
        it = kolom(header, ["wetbulb"])
        ig = kolom(header, ["globe"])
        ia = kolom(header, ["temperature"])

        nilai = []
        selisih = []

        for row in baris[1:]:
            if None in (iw, it, ig, ia):
                continue

            wbgt = angka(row[iw])
            tw = angka(row[it])
            tg = angka(row[ig])
            ta = angka(row[ia])

            if None not in (wbgt, tw, tg, ta):
                rumus = (
                    0.7 * tw + 0.2 * tg + 0.1 * ta
                    if mode == "outdoor"
                    else 0.7 * tw + 0.3 * tg
                )
                nilai.append(wbgt)
                selisih.append(abs(wbgt - rumus))

        print("\n", nama)
        print("Jumlah:", len(nilai))
        print(f"Rentang WBGT: {min(nilai):.3f}–{max(nilai):.3f}")
        print(f"MAE terhadap rumus WBGT: {statistics.mean(selisih):.6f} °C")

    wb.close()

if __name__ == "__main__":
    utama()
