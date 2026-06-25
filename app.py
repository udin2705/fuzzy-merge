import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Merge Excel - Nama 100%", layout="wide")

st.title("🗂️ Penggabung Excel (Join 'No Rek' & 'Nama' 100% Sama)")
st.write("Skrip ini hanya akan menggabungkan data ke samping jika **No Rek** cocok DAN kolom **Nama** sama persis (100%).")

# Input nama kolom kunci dasar
key_rek = st.text_input("Nama kolom Rekening:", value="No Rek")
key_nama = st.text_input("Nama kolom Nama:", value="Nama")

# Pilihan metode join data
join_method = st.radio(
    "Pilih Metode Penggabungan Utama:",
    options=[
        "Tampilkan semua rekening dari sheet pertama/acuan (Left Join)", 
        "Hanya tampilkan jika cocok di semua sheet (Inner Join)"
    ],
    index=0
)
how_strategy = "left" if "Left" in join_method else "inner"

# 1. Fitur Unggah File
uploaded_files = st.file_uploader(
    "Unggah File Excel Anda (.xlsx)", 
    accept_multiple_files=True, 
    type=["xlsx"]
)

if uploaded_files:
    all_dfs = []
    
    # 2. Proses membaca setiap sheet dari semua file
    for uploaded_file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                # Bersihkan spasi di nama kolom
                df.columns = df.columns.str.strip()
                
                if key_rek in df.columns and key_nama in df.columns:
                    # Standarisasi No Rek (ubah ke string dan hapus spasi)
                    df[key_rek] = df[key_rek].astype(str).str.strip()
                    
                    # Standarisasi Nama (hapus spasi ujung dan ubah ke huruf kapital untuk pencocokan 100% yang adil)
                    df[key_nama] = df[key_nama].astype(str).str.strip().str.upper()
                    
                    # Beri penanda unik pada kolom informasi lainnya agar tidak bentrok
                    suffix = f"_{sheet_name}"
                    rename_dict = {
                        col: f"{col}{suffix}" for col in df.columns 
                        if col != key_rek and col != key_nama
                    }
                    df = df.rename(columns=rename_dict)
                    
                    all_dfs.append((sheet_name, df))
                else:
                    st.warning(f"⚠️ Kolom '{key_rek}' atau '{key_nama}' tidak ditemukan di Sheet: **{sheet_name}** (Dilewati)")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    # 3. Proses Join Berdasarkan Multi-Kolom (No Rek & Nama)
    if all_dfs:
        # Gunakan sheet pertama sebagai dasar basis data
        base_sheet_name, merged_df = all_dfs[0]
        
        # Lakukan perulangan untuk melakukan join dengan sheet berikutnya
        for next_sheet_name, next_df in all_dfs[1:]:
            # Join didasarkan pada dua kolom sekaligus: ['No Rek', 'Nama']
            merged_df = pd.merge(
                merged_df, 
                next_df, 
                on=[key_rek, key_nama], 
                how=how_strategy
            )
        
        st.success(f"🎉 Berhasil memproses gabungan {len(all_dfs)} sheet dengan validasi nama 100%! Total hasil: {len(merged_df)} baris.")
        
        # Tampilkan pratinjau data hasil gabungan
        with st.expander("👁️ Pratinjau Hasil Gabungan (Satu Kolom Nama Acuan)"):
            st.dataframe(merged_df.head(10))
            
        # 4. Unduh Hasil Akhir
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='Hasil_Join_100_Persen')
        
        excel_bytes = buffer.getvalue()
        
        st.download_button(
            label="⬇️ Unduh File Excel Hasil Join",
            data=excel_bytes,
            file_name="merge_rek_dan_nama_100.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
