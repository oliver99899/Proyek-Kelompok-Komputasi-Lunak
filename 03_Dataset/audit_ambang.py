
import zipfile
import io
from pathlib import Path
from openpyxl import load_workbook

lokasi = Path(input("Path ZIP: ").strip().strip('"'))

with zipfile.ZipFile(lokasi) as zi:
    nama = next(x for x in zi.namelist() if x.endswith("Dataset.xlsx"))
    wb = load_workbook(io.BytesIO(zi.read(nama)), data_only=True)

for sheet in ["Final Lab Test 1", "Final Lab Test 2", "Final Lab Test 3"]:
    ws = wb[sheet]
    data = list(ws.iter_rows(values_only=True))
    header = [str(x).strip() for x in data[0]]

    ir = header.index("WBGT Reference")
    ip = header.index("WBGT Developed System")

    pasang = [
        (float(row[ir]), float(row[ip]))
        for row in data[1:]
        if row[ir] is not None and row[ip] is not None
    ]

    print("\n" + "=" * 70)
    print(sheet)
    print("=" * 70)

    for ambang in [27.5, 28, 29, 30, 31, 32]:
        positif = [
            (i, r, p)
            for i, (r, p) in enumerate(pasang, 2)
            if r >= ambang
        ]

        kurang = [
            (i, r, p)
            for i, r, p in positif
            if p < ambang
        ]

        lebih = [
            (i, r, p)
            for i, (r, p) in enumerate(pasang, 2)
            if r < ambang and p >= ambang
        ]

        n = len(positif)
        persen = 100 * len(kurang) / n if n else 0

        print(f"\nAmbang {ambang:.1f} °C")
        print("Referensi mencapai ambang:", n)
        print("Prototipe di bawah ambang:", len(kurang))
        print("Potensi underclassification:", f"{persen:.2f}%")
        print("Potensi overclassification:", len(lebih))

        if kurang:
            print("Contoh baris (Excel, Referensi, Prototipe):")
            for item in kurang[:5]:
                print(item)

wb.close()
