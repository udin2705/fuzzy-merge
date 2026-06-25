import streamlit as st
import pandas as pd
import io

# Konfigurasi halaman utama
st.set_page_config(page_title="Merge Excel No Rek", layout="wide")

st.title("🗂️ Penggabung Excel (Join Berdasarkan 'No Rek')")
st.write("Unggah beberapa file atau sheet Excel untuk digabungkan ke samping (horizontal) secara otomatis.")

# Input nama kolom kunci (Default: No Rek)
key_column = st.text_input("Nama kolom kunci untuk dasar join:", value="No Rek")

# Pilihan tipe join data
join_method = st.radio(
    "Pilih Metode Penggabungan Data:",
    options=[
        "Tampilkan SEMUA nomor rekening dari seluruh file/sheet (Outer Join)", 
        "Hanya tampilkan nomor rekening yang COCOK di semua file/sheet (Inner Join)"
    ],
    index=0
)
how_strategy = "outer" if "SEMUA" in join_method else "inner"

# 1. Fitur Unggah File (Mendukung banyak file sekaligus)
uploaded_files = st.file_uploader(
    "Unggah File Excel Anda (.xlsx)", 
    accept_multiple_files=True, 
    type=["xlsx"]
)

if uploaded_files:
    all_dfs = []
    
    # 2. Proses membaca setiap sheet dari semua file yang diunggah
    for uploaded_file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                # Bersihkan spasi di nama kolom
                df.columns = df.columns.str.strip()
                
                if key_column in df.columns:
                    # Bersihkan spasi dan ubah tipe data kolom kunci menjadi teks (String)
                    df[key_column] = df[key_column].astype(str).str.strip()
                    
                    # Beri akhiran unik pada kolom lain agar tahu asal sumber data sheet-nya
                    suffix = f"_{uploaded_file.name.split('.')[0]}_{sheet_name}"
                    rename_dict = {col: f"{col}{suffix}" for col in df.columns if col != key_column}
                    df = df.rename(columns=rename_dict)
                    
                    all_dfs.append(df)
                else:
                    st.warning(f"⚠️ Kolom '{key_column}' tidak ditemukan di File: **{uploaded_file.name}** | Sheet: **{sheet_name}** (Dilewati)")
        except Exception as e:
            st.error(f"Gagal membaca file {uploaded_file.name}: {e}")

    # 3. Eksekusi Proses Join Penggabungan Data
    if all_dfs:
        # Gunakan dataframe pertama sebagai dasar awal
        merged_df = all_dfs[0]
        
        # Gabungkan secara berurutan dengan dataframe berikutnya
        for next_df in all_dfs[1:]:
            merged_df = pd.merge(merged_df, next_df, on=key_column, how=how_strategy)
        
        st.success(f"🎉 Berhasil menggabungkan {len(all_dfs)} data tabel! Total: {len(merged_df)} baris.")
        
        # Tampilkan pratinjau hasil gabungan di web streamlit
        with st.expander("👁️ Pratinjau Hasil Gabungan Kolom"):
            st.dataframe(merged_df.head(10))
            
        # 4. Fitur Ekspor / Unduh File Hasil Merge
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='Hasil_Merge')
        
        excel_bytes = buffer.getvalue()
        
        st.download_button(
            label="⬇️ Unduh File Excel Hasil Merge",
            data=excel_bytes,
            file_name="excel_merge_no_rek.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Silakan unggah dokumen yang memiliki struktur kolom yang sesuai untuk memulai.")
