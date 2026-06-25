import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher

st.set_page_config(page_title="Merge Excel Custom Similarity", layout="wide")

st.title("🗂️ Penggabung Excel dengan Input Persentase Kemiripan Nama")
st.write("Tentukan batas kemiripan nama sendiri. Jika kemiripan di atas target, kolom nama akan otomatis menyatu.")

# 1. Parameter Input di Sidebar / Halaman Utama
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("⚙️ Pengaturan Gabung")
    key_column = "No Rek"
    name_column = "Nama"
    
    # Input interaktif untuk menentukan persentase kemiripan nama
    similarity_threshold = st.slider(
        "Batas Minimal Kemiripan Nama (%):", 
        min_value=0, 
        max_value=100, 
        value=90, 
        step=5,
        help="Jika kemiripan nama di antar sheet mencapai angka ini atau lebih, kolom nama akan digabung menjadi satu."
    )
    
    join_method = st.radio(
        "Metode Penggabungan Nomor Rekening:",
        options=[
            "Tampilkan SEMUA nomor rekening (Outer Join)", 
            "Hanya nomor rekening yang cocok di semua sheet (Inner Join)"
        ],
        index=0
    )
    how_strategy = "outer" if "SEMUA" in join_method else "inner"

# Fungsi untuk menghitung persentase kemiripan teks (Fuzzy Match)
def hitung_persen_kemiripan(teks1, teks2):
    if pd.isna(teks1) or pd.isna(teks2):
        return 0.0
    return SequenceMatcher(None, str(teks1).strip().lower(), str(teks2).strip().lower()).ratio() * 100

# 2. Upload File
with col2:
    st.subheader("📁 Unggah Dokumen")
    uploaded_files = st.file_uploader("Unggah File Excel Anda (.xlsx)", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    all_dfs = []
    
    # Membaca seluruh file dan sheet yang diunggah
    for uploaded_file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                df.columns = df.columns.str.strip()
                
                if key_column in df.columns:
                    df[key_column] = df[key_column].astype(str).str.strip()
                    if name_column in df.columns:
                        df[name_column] = df[name_column].astype(str).str.strip()
                    
                    # Simpan informasi asal nama sheet
                    df.attrs['sheet_name'] = sheet_name
                    all_dfs.append(df)
                else:
                    st.warning(f"⚠️ Kolom '{key_column}' tidak ditemukan di Sheet: **{sheet_name}** (Dilewati)")
        except Exception as e:
            st.error(f"Gagal membaca file {uploaded_file.name}: {e}")

    # 3. Proses Penggabungan dengan Aturan Persentase Kustom
    if all_dfs:
        # Gunakan sheet pertama sebagai basis acuan utama
        merged_df = all_dfs[0].copy()
        
        for next_df in all_dfs[1:]:
            sheet_name = next_df.attrs['sheet_name']
            
            if name_column in merged_df.columns and name_column in next_df.columns:
                # Lakukan merge temporary untuk menghitung skor kecocokan nama per baris
                temp_merge = pd.merge(
                    merged_df[[key_column, name_column]], 
                    next_df[[key_column, name_column]], 
                    on=key_column, 
                    how=how_strategy, 
                    suffixes=('', '_temp')
                )
                
                # Hitung skor kemiripan untuk setiap baris data
                temp_merge['skor_cocok'] = temp_merge.apply(
                    lambda r: hitung_persen_kemiripan(r[name_column], r[f"{name_column}_temp"]), axis=1
                )
                
                # Filter baris yang nilai kemiripannya di bawah instruksi user (Gagal Match)
                gagal_match = temp_merge[temp_merge['skor_cocok'] < similarity_threshold]
                
                # Mapping No Rek yang namanya dianggap tidak mirip
                rename_map = dict(zip(gagal_match[key_column], gagal_match[f"{name_column}_temp"]))
                
                # Pindahkan nama yang tidak mirip ke kolom baru (Nama_NamaSheet)
                next_df[f"{name_column}_{sheet_name}"] = next_df.apply(
                    lambda row: row[name_column] if row[key_column] in rename_map else None, axis=1
                )
                
                # Kosongkan kolom Nama utama pada baris yang tidak mirip agar tidak menimpa kolom utama saat di-merge
                next_df.loc[next_df[key_column].isin(rename_map.keys()), name_column] = None

            # Beri suffix unik untuk kolom lain selain kunci utama
            rename_dict = {}
            for col in next_df.columns:
                if col != key_column and col != name_column and not col.endswith(f"_{sheet_name}"):
                    rename_dict[col] = f"{col}_{sheet_name}"
            
            next_df = next_df.rename(columns=rename_dict)
            
            # Gabungkan tabel secara horizontal
            merged_df = pd.merge(merged_df, next_df, on=key_column, how=how_strategy)
            
            # Bersihkan dan satukan kembali kolom Nama sisa pecahan merge otomatis Pandas jika ada
            if f"{name_column}_x" in merged_df.columns and f"{name_column}_y" in merged_df.columns:
                merged_df[name_column] = merged_df[f"{name_column}_x"].fillna(merged_df[f"{name_column}_y"])
                merged_df = merged_df.drop(columns=[f"{name_column}_x", f"{name_column}_y"])

        # Menampilkan output hasil pemrosesan
        st.success(f"🎉 Selesai memproses! Total data gabungan: {len(merged_df)} baris.")
        
        with st.expander("👁️ Tampilkan Sampel Data Gabungan"):
            st.dataframe(merged_df.head(10))
            
        # 4. Pembuatan Fitur Download Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='Hasil_Merge_Custom')
        
        excel_bytes = buffer.getvalue()
        
        st.download_button(
            label="⬇️ Unduh Hasil Analisis Excel",
            data=excel_bytes,
            file_name="merge_excel_similarity_custom.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
