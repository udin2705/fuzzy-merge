import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gabung Sheet Excel", layout="centered")

st.title("🗂️ Penggabung Sheet/File Excel")
st.write("Unggah file Excel Anda dan gabungkan semua data ke dalam satu file unduhan.")

# 1. Upload File
uploaded_files = st.file_uploader("Pilih file Excel (.xlsx)", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    all_data = []
    
    # 2. Proses membaca sheet dari setiap file yang diupload
    for file in uploaded_files:
        try:
            # Membaca semua sheet di dalam satu file Excel
            excel_file = pd.ExcelFile(file)
            sheet_names = excel_file.sheet_names
            
            for sheet in sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet)
                
                # Tambahkan kolom informasi sumber jika diperlukan
                df['Nama File'] = file.name
                df['Nama Sheet'] = sheet
                
                all_data.append(df)
                
        except Exception as e:
            st.error(f"Gagal membaca file {file.name}: {e}")

    if all_data:
        # Menggabungkan semua data secara vertikal (concat)
        master_df = pd.concat(all_data, ignore_index=True)
        
        st.success(f"Berhasil menggabungkan {len(all_data)} sheet! Total baris: {len(master_df)}")
        
        # Pratinjau Data
        with st.expander("👁️ Lihat Pratinjau Data Gabungan"):
            st.dataframe(master_df.head(10))
            
        # 3. Tombol Download Hasil
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            master_df.to_excel(writer, index=False, sheet_name='Hasil_Gabungan')
        
        excel_data = buffer.getvalue()
        
        st.download_button(
            label="⬇️ Unduh File Excel Gabungan",
            data=excel_data,
            file_name="data_gabungan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
