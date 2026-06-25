import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from thefuzz import fuzz

# 1. BACA KEDUA FILE EXCEL
file1_path = "file_pertama.xlsx"
file2_path = "file_kedua.xlsx"
output_path = "hasil_analisis_gaji.xlsx"

df1 = pd.read_excel(file1_path)  # Kolom: Nama Karyawan, No. Rekening, Gaji
df2 = pd.read_excel(
    file2_path
)  # Kolom: Nama Karyawan, No. Rekening, Potongan Angsuran

# Pastikan tipe data No. Rekening berupa string dan bersih dari spasi
df1["No. Rekening"] = df1["No. Rekening"].astype(str).str.strip()
df2["No. Rekening"] = df2["No. Rekening"].astype(str).str.strip()

# Buat dictionary dari file kedua untuk mempercepat pencarian berdasarkan No. Rekening
# Struktur: { 'no_rekening': ('nama_di_file2', potongan_angsuran) }
dict_file2 = {}
for _, row in df2.iterrows():
    dict_file2[str(row["No. Rekening"])] = (
        str(row["Nama Karyawan"]),
        row["Potongan Angsuran"],
    )

# 2. PROSES PENCOCOKAN & HITUNG SIMILARITY
nama_baru_list = []
potongan_list = []
similarity_list = []

for _, row in df1.iterrows():
    nama1 = str(row["Nama Karyawan"]).strip()
    norek = str(row["No. Rekening"]).strip()

    # Cek apakah No. Rekening ada di file kedua
    if norek in dict_file2:
        nama2, potongan = dict_file2[norek]
        nama2 = nama2.strip()

        # Hitung skor kemiripan antara nama di file 1 dan file 2 (skala 0 - 100)
        skor = fuzz.ratio(nama1.lower(), nama2.lower())

        potongan_list.append(potongan)
        similarity_list.append(skor)
    else:
        # Jika No. Rekening tidak ditemukan di file kedua
        potongan_list.append(0)
        similarity_list.append(
            0
        )  # Dianggap 0% karena data orangnya tidak ada di file 2

# Tambahkan kolom baru ke DataFrame pertama (tetap menggunakan Nama Karyawan asli)
df1["Potongan Angsuran"] = potongan_list
df1["Potongan Angsuran"] = df1["Potongan Angsuran"].fillna(0)
df1["Similarity_Score"] = similarity_list

# Simpan hasil sementara ke Excel
df1.to_excel(output_path, index=False)

# 3. PROSES MEWARNAI CELL MENGGUNAKAN OPENPYXL
wb = load_workbook(output_path)
ws = wb.active

# Tentukan warna (Hex Code)
warna_hijau = PatternFill(
    start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
)  # Hijau Lembut
warna_kuning = PatternFill(
    start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
)  # Kuning Lembut
warna_merah = PatternFill(
    start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
)  # Merah Lembut

# Cari tahu posisi kolom "Nama Karyawan" dan "Similarity_Score"
idx_nama = None
idx_score = None

for col in range(1, ws.max_column + 1):
    header = ws.cell(row=1, column=col).value
    if header == "Nama Karyawan":
        idx_nama = col
    elif header == "Similarity_Score":
        idx_score = col

# Lakukan pewarnaan pada kolom Nama Karyawan berdasarkan nilai Similarity_Score
if idx_nama and idx_score:
    for row in range(2, ws.max_row + 1):
        skor_cell = ws.cell(row=row, column=idx_score).value
        nama_cell = ws.cell(row=row, column=idx_nama)

        if skor_cell > 80:
            nama_cell.fill = warna_hijau
        elif 30 <= skor_cell <= 80:
            nama_cell.fill = warna_kuning
        else:  # di bawah 30%
            nama_cell.fill = warna_merah

# Hapus kolom bantuan 'Similarity_Score' agar file akhir terlihat rapi
ws.delete_cols(idx_score)

# Simpan file Excel yang sudah diwarnai
wb.save(output_path)
print(f"Selesai! File berhasil digabungkan dan diwarnai di: {output_path}")
