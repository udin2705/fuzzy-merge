import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl

st.set_page_config(page_title="Pencocokan Nama Excel", layout="wide")

st.title("📑 Konsolidasi Kolom 'Nama' Berdasarkan Acuan Sheet1")
st.write("Skrip ini mencocokkan 'No Rek', mempertahankan kolom 'Nama' Sheet1, dan mewarnai kuning nama yang berbeda < 90% di sheet lain.")

# Fungsi untuk menghitung persentase kemiripan teks
def hitung_kemiripan(teks1, teks2):
    if pd.isna(teks1) or pd.isna(teks2):
        return 0.0
    return SequenceMatcher(None, str(teks1).strip().lower(), str(teks2).strip().lower()).ratio() * 100

# 1. Unggah File
uploaded_files = st.file_uploader("Unggah file Excel (.xlsx)", accept_multiple_files=False, type=["xlsx"])

if uploaded_files:
    excel_file = pd.ExcelFile(uploaded_files)
    sheet_names = excel_file.sheet_names
    
    if "Sheet1" not in sheet_names:
        st.error("❌ Eror: Aplikasi mewajibkan adanya sheet bernama 'Sheet1' sebagai acuan utama.")
    else:
        # Membaca Sheet1 sebagai acuan dasar
        df_master = pd.read_excel(excel_file, sheet_name="Sheet1")
        df_master.columns = df_master.columns.str.strip()
        
        # Validasi kolom wajib di Sheet1
        if "No Rek" not in df_master.columns or "Nama" not in df_master.columns:
            st.error("❌ Eror: 'Sheet1' harus memiliki kolom 'No Rek' dan 'Nama'.")
        else:
            # Standarisasi kolom kunci
            df_master["No Rek"] = df_master["No Rek"].astype(str).str.strip()
            
            # List untuk menyimpan informasi kolom kemiripan yang perlu diwarnai nanti
            kolom_skor_kemiripan = []
            
            # 2. Gabungkan sheet lain ke Sheet1 berdasarkan No Rek
            for sheet in sheet_names:
                if sheet == "Sheet1":
                    continue
                    
                df_sheet = pd.read_excel(excel_file, sheet_name=sheet)
                df_sheet.columns = df_sheet.columns.str.strip()
                
                if "No Rek" in df_sheet.columns and "Nama" in df_sheet.columns:
                    df_sheet["No Rek"] = df_sheet["No Rek"].astype(str).str.strip()
                    
                    # Ambil hanya No Rek dan Nama dari sheet pembanding
                    df_temp = df_sheet[["No Rek", "Nama"]].rename(columns={"Nama": f"Nama_{sheet}"})
                    
                    # Hubungkan ke master data (Left Join agar Sheet1 tetap jadi acuan baris)
                    df_master = pd.merge(df_master, df_temp, on="No Rek", how="left")
                    
                    # Hitung skor kemiripan antara Nama (Sheet1) dan Nama_(Sheet)
                    nama_skor_col = f"Kemiripan_Nama_{sheet} (%)"
                    df_master[nama_skor_col] = df_master.apply(
                        lambda row: round(hitung_kemiripan(row["Nama"], row[f"Nama_{sheet}"]), 1), axis=1
                    )
                    
                    kolom_skor_kemiripan.append((f"Nama_{sheet}", nama_skor_col))
                else:
                    st.warning(f"⚠️ Sheet '{sheet}' dilewati karena tidak memiliki kolom 'No Rek' atau 'Nama'.")
            
            # 3. Pratinjau Sementara di Aplikasi
            st.success("🎉 Data berhasil dicocokkan! Memproses pemformatan file...")
            
            # Buat workbook openpyxl baru untuk menerapkan warna secara akurat
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Hasil_Gabungan"
            
            # Tulis DataFrame ke sheet openpyxl
            for r in dataframe_to_rows(df_master, index=False, header=True):
                ws.append(r)
                
            # Dapatkan indeks posisi kolom untuk proses pewarnaan cell
            headers = [cell.value for cell in ws[1]]
            fill_kuning = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            
            # 4. Pewarnaan Cell Berdasarkan Logika Kriteria < 90%
            for nama_col, skor_col in kolom_skor_kemiripan:
                idx_nama = headers.index(nama_col) + 1
                idx_skor = headers.index(skor_col) + 1
                
                # Iterasi baris data (mulai dari baris ke-2 untuk melewati header)
                for row in range(2, ws.max_row + 1):
                    skor_val = ws.cell(row=row, column=idx_skor).value
                    # Jika skor di bawah 90%, beri warna kuning pada sel nama sheet tersebut
                    if skor_val is not None and skor_val < 90.0:
                        ws.cell(row=row, column=idx_nama).fill = fill_kuning
            
            # 5. Sembunyikan/Hapus Kolom Nama Duplikat dari Tampilan (Opsional: Tetap disimpan tapi disembunyikan agar user bisa audit)
            # Di sini kita biarkan kolom bantu tersebut agar warna kuningnya terlihat jelas saat file dibuka.
            
            # Konversi hasil ke objek bytes agar siap diunduh
            buffer = io.BytesIO()
            wb.save(buffer)
            excel_data = buffer.getvalue()
            
            # Tampilkan 10 baris pertama di layar web sebagai sampel
            st.dataframe(df_master.head(10))
            
            st.download_button(
                label="⬇️ Unduh File Excel Gabungan Berwarna (Kuning < 90%)",
                data=excel_data,
                file_name="analisis_nama_gabungan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
