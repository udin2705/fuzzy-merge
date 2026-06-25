import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Pengaturan halaman website
st.set_page_config(
    page_title="Multi-Excel Merger with Preview", page_icon="📑", layout="wide"
)

st.title("📑 Multi-Excel Advanced Merger & Preview")
st.write(
    "Unggah beberapa file Excel sekaligus. Sistem akan menggabungkan data, mendeteksi kolom baru, "
    "menampilkan pratinjau hasil, dan mewarnai data berdasarkan kemiripan struktur kolom dengan file pertama."
)

# 1. ANTARMUKA UNTUK UNGGAH BANYAK FILE
uploaded_files = st.file_uploader(
    "Upload Semua File Excel Anda (Bisa pilih banyak sekaligus)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) < 2:
        st.warning("⚠️ Silahkan unggah minimal 2 file Excel untuk digabungkan.")
    else:
        st.info(
            f"💡 {len(uploaded_files)} file terdeteksi. File pertama ({uploaded_files[0].name}) akan digunakan sebagai acuan utama."
        )

        if st.button("🚀 Mulai Gabungkan & Analisis", use_container_width=True):
            try:
                # 2. BACA FILE PERTAMA SEBAGAI ACUAN UTAMA
                file_acuan = uploaded_files[0]
                df_master = pd.read_excel(file_acuan)

                # Format nama kolom file acuan agar bersih
                df_master.columns = [str(c).strip() for c in df_master.columns]
                kolom_acuan_asli = list(df_master.columns)

                # Tambahkan kolom penanda warna: 'H' = Hijau, 'M' = Merah
                df_master["_row_color_flag"] = "H"

                # 3. PROSES FILE KEDUA DAN SETERUSNYA
                for file_tambahan in uploaded_files[1:]:
                    df_temp = pd.read_excel(file_tambahan)
                    df_temp.columns = [str(c).strip() for c in df_temp.columns]

                    # Hitung kemiripan (similarity) set kolom
                    kolom_master_set = set(kolom_acuan_asli)
                    kolom_temp_set = set(df_temp.columns)
                    kolom_sama = kolom_master_set.intersection(kolom_temp_set)

                    total_kolom_gabungan = len(
                        kolom_master_set.union(kolom_temp_set)
                    )
                    skor_kemiripan_kolom = (
                        (len(kolom_sama) / total_kolom_gabungan) * 100
                        if total_kolom_gabungan > 0
                        else 0
                    )

                    # Tentukan penanda warna baris
                    if skor_kemiripan_kolom >= 80:
                        df_temp["_row_color_flag"] = "H"
                    else:
                        df_temp["_row_color_flag"] = "M"

                    # Gabungkan data
                    df_master = pd.concat([df_master, df_temp], ignore_index=True)

                # Isi nilai kosong (NaN) dengan string kosong
                df_master = df_master.fillna("")

                # Susun ulang urutan kolom
                semua_kolom = list(df_master.columns)
                semua_kolom.remove("_row_color_flag")
                kolom_baru = [c for c in semua_kolom if c not in kolom_acuan_asli]
                urutan_kolom_final = (
                    kolom_acuan_asli + kolom_baru + ["_row_color_flag"]
                )
                df_master = df_master[urutan_kolom_final]

                # 4. FITUR PRATINJAU TABEL (PREVIEW DATA) DI WEB
                st.subheader("👀 Pratinjau Tabel Hasil Gabungan")
                st.write(
                    "Berikut adalah tampilan sementara data Anda. Baris dengan latar belakang "
                    "hijau/merah menandakan tingkat kecocokan kolom."
                )

                # Fungsi styling untuk pratinjau tabel Streamlit
                def style_row(row):
                    # Ambil flag warna di kolom terakhir
                    flag = row["_row_color_flag"]
                    if flag == "H":
                        return ["background-color: #C6EFCE; color: #006100"] * len(
                            row
                        )
                    elif flag == "M":
                        return ["background-color: #FFC7CE; color: #9C0006"] * len(
                            row
                        )
                    return [""] * len(row)

                # Tampilkan dataframe dengan style warna (kolom flag disembunyikan di preview)
                df_preview = df_master.copy()
                st.dataframe(
                    df_preview.style.apply(style_row, axis=1).hide(
                        axis="columns", subset=["_row_color_flag"]
                    ),
                    use_container_width=True,
                )

                # 5. LOGIKA PEWARNAAN FILE EXCEL ASLI (OPENPYXL)
                buffer = io.BytesIO()
                df_master.to_excel(buffer, index=False)
                buffer.seek(0)

                wb = load_workbook(buffer)
                ws = wb.active

                warna_hijau = PatternFill(
                    start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
                )
                warna_merah = PatternFill(
                    start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
                )

                idx_flag = ws.max_column

                for row in range(2, ws.max_row + 1):
                    flag_val = ws.cell(row=row, column=idx_flag).value
                    if flag_val == "H":
                        for col in range(1, idx_flag):
                            ws.cell(row=row, column=col).fill = warna_hijau
                    elif flag_val == "M":
                        for col in range(1, idx_flag):
                            ws.cell(row=row, column=col).fill = warna_merah

                # Hapus kolom bantuan flag agar file hasil bersih
                ws.delete_cols(idx_flag)

                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)

                st.success(
                    "🎉 Analisis selesai! Struktur warna di atas akan sama persis dengan file Excel yang diunduh."
                )

                # 6. TOMBOL DOWNLOAD HASIL AKHIR
                st.download_button(
                    label="📥 Unduh File Gabungan (.xlsx)",
                    data=output_buffer,
                    file_name="gabungan_excel_with_preview.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")
