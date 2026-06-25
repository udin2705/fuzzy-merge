import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from thefuzz import process, fuzz

st.set_page_config(
    page_title="Rekening Fuzzy Merger",
    page_icon="💳",
    layout="wide",
)

st.title("💳 No. Rekening Fuzzy Merger & Validator")
st.write(
    "Unggah beberapa file Excel. Aplikasi akan melakukan **Fuzzy Matching pada No. Rekening** yang typo/salah ketik "
    "agar otomatis diarahkan ke **No. Rekening Valid** dari file acuan pertama dengan validasi kesesuaian nama."
)

uploaded_files = st.file_uploader(
    "Upload Semua File Excel Anda (Minimal 2 file)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) < 2:
        st.warning("⚠️ Silahkan unggah minimal 2 file Excel untuk digabungkan.")
    else:
        st.info(
            f"💡 File acuan utama: **{uploaded_files.name}**. Nomor rekening di file ini dianggap sebagai **Rekening Valid** yang sah."
        )

        if st.button("🚀 Proses & Validasi Nomor Rekening", use_container_width=True):
            try:
                # 1. BACA FILE ACUAN UTAMA (REKENING VALID)
                df_master = pd.read_excel(uploaded_files)
                df_master.columns = [str(c).strip() for c in df_master.columns]
                
                if "Nama Karyawan" not in df_master.columns or "No. Rekening" not in df_master.columns:
                    st.error("File pertama wajib memiliki kolom 'Nama Karyawan' dan 'No. Rekening'!")
                    st.stop()
                
                df_master["No. Rekening"] = df_master["No. Rekening"].astype(str).str.strip()
                list_rek_valid = df_master["No. Rekening"].tolist()
                
                # Buat dictionary master untuk verifikasi silang { 'rek_valid': 'nama_karyawan' }
                dict_master_user = dict(zip(df_master["No. Rekening"], df_master["Nama Karyawan"]))
                
                # List penanda warna khusus untuk kolom 'No. Rekening' di file master hasil akhir
                warna_rek_list = ["H"] * len(df_master)

                # 2. PROSES FILE TAMBAHAN (FILE POTONGAN)
                for file_tambahan in uploaded_files[1:]:
                    df_temp = pd.read_excel(file_tambahan)
                    df_temp.columns = [str(c).strip() for c in df_temp.columns]
                    
                    if "Nama Karyawan" not in df_temp.columns or "No. Rekening" not in df_temp.columns:
                        st.error(f"File '{file_tambahan.name}' wajib memiliki kolom 'Nama Karyawan' dan 'No. Rekening'!")
                        st.stop()
                        
                    df_temp["No. Rekening"] = df_temp["No. Rekening"].astype(str).str.strip()

                    # Inisialisasi wadah untuk kolom baru jika ada dari file tambahan
                    for col in df_temp.columns:
                        if col not in df_master.columns and col != "Nama Karyawan" and col != "No. Rekening":
                            df_master[col] = 0.0

                    # Iterasi setiap data potongan di file tambahan untuk dicarikan rekening validnya
                    for _, row_temp in df_temp.iterrows():
                        rek_temp = str(row_temp["No. Rekening"]).strip()
                        nama_temp = str(row_temp["Nama Karyawan"]).strip()
                        
                        # FUZZY MATCHING: Cari nomor rekening paling mirip di daftar rekening master yang valid
                        rek_valid_terdekat, skor_rek = process.extractOne(rek_temp, list_rek_valid)
                        
                        # VALIDASI GANDA: Cek apakah nama di file tambahan mirip dengan nama pemilik asli rekening tersebut
                        nama_pemilik_asli = dict_master_user.get(rek_valid_terdekat, "")
                        skor_nama = fuzz.ratio(nama_temp.lower(), str(nama_pemilik_asli).lower())
                        
                        # Hitung skor total kombinasi (Bobot: 60% Rekening, 40% Kesesuaian Nama)
                        skor_final = (skor_rek * 0.6) + (skor_nama * 0.4)
                        
                        # Tentukan penanda warna berdasarkan kriteria kecocokan final
                        if skor_final > 80:
                            flag_warna = "H"  # Hijau (Aman, diarahkan ke rekening valid)
                        elif 30 <= skor_final <= 80:
                            flag_warna = "K"  # Kuning (Ragu-ragu, sistem tetap memasukkan tapi tandai kuning)
                        else:
                            flag_warna = "M"  # Merah (Sangat tidak cocok / data tidak ditemukan)
                            
                        # Cari indeks baris di df_master yang memiliki rekening valid terpilih tersebut
                        idx_master = df_master[df_master["No. Rekening"] == rek_valid_terdekat].index
                        
                        if not idx_master.empty:
                            target_idx = idx_master[0]
                            # Masukkan nilai potongan/data lain ke baris rekening valid tersebut
                            for col in df_temp.columns:
                                if col != "Nama Karyawan" and col != "No. Rekening":
                                    val_baru = row_temp[col]
                                    # Jika kolom berupa angka/keuangan, kita akumulasikan nilainya
                                    try:
                                        val_lama = df_master.at[target_idx, col]
                                        df_master.at[target_idx, col] = float(val_lama if val_lama != "" else 0) + float(val_baru)
                                    except:
                                        df_master.at[target_idx, col] = val_baru
                                        
                            # Simpan status warna terburuk yang dialami rekening ini jika terjadi multi-proses
                            if warna_rek_list[target_idx] == "H" or flag_warna == "M":
                                warna_rek_list[target_idx] = flag_warna

                # Isi nilai kosong (NaN) dengan string kosong atau angka 0 agar rapi
                df_master = df_master.fillna("")

                # 3. PRATINJAU TABEL DI WEBSITE
                st.subheader("👀 Pratinjau Hasil Validasi Rekening")
                st.write("Warna di bawah ini diterapkan pada kolom **No. Rekening** untuk mengukur validitas data potongan yang masuk.")

                def style_rekening_column(df_input):
                    df_style = pd.DataFrame("", index=df_input.index, columns=df_input.columns)
                    for r in range(len(df_input)):
                        flag = warna_rek_list[r]
                        if flag == "H":
                            df_style.loc[df_input.index[r], "No. Rekening"] = "background-color: #C6EFCE; color: #006100"
                        elif flag == "K":
                            df_style.loc[df_input.index[r], "No. Rekening"] = "background-color: #FFEB9C; color: #9C0006"
                        elif flag == "M":
                            df_style.loc[df_input.index[r], "No. Rekening"] = "background-color: #FFC7CE; color: #9C0006"
                    return df_style

                st.dataframe(df_master.style.apply(style_rekening_column, axis=None), use_container_width=True)

                # 4. PROSES PEWARNAAN CELL EXCEL ASLI (OPENPYXL)
                buffer = io.BytesIO()
                df_master.to_excel(buffer, index=False)
                buffer.seek(0)

                wb = load_workbook(buffer)
                ws = wb.active

                warna_hijau = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                warna_kuning = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                warna_merah = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

                # Cari indeks kolom 'No. Rekening' di file excel hasil
                idx_rek = next(col for col in range(1, ws.max_column + 1) if ws.cell(row=1, column=col).value == "No. Rekening")

                # Warnai cell No. Rekening berdasarkan list warna_rek_list
                for r_idx in range(2, ws.max_row + 1):
                    flag_val = warna_rek_list[r_idx - 2]
                    cell = ws.cell(row=r_idx, column=idx_rek)

                    if flag_val == "H":
                        cell.fill = warna_hijau
                    elif flag_val == "K":
                        cell.fill = warna_kuning
                    elif flag_val == "M":
                        cell.fill = warna_merah

                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)

                st.success("🎉 Berhasil merujuk rekening typo ke nomor rekening valid!")

                # 5. TOMBOL DOWNLOAD
                st.download_button(
                    label="📥 Unduh File Hasil Validasi (.xlsx)",
                    data=output_buffer,
                    file_name="gabungan_rekening_valid_fuzzy.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses data: {e}")
