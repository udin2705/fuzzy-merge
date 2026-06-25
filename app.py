import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher

st.set_page_config(page_title="Merge No Rek & Similarity Nama", layout="wide")

st.title("🗂️ Penggabung Excel (Validasi Nama per No Rek)")
st.write("Skrip ini mencocokkan **No Rek**. Jika baris No Rek yang sama memiliki **Nama** dengan kemiripan di bawah target slider, nama sheet pembanding akan dipecah ke kolom kanan baru.")

# 1. Konfigurasi Kontrol di Sidebar / Halaman Utama
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("⚙️ Atur Batas Toleransi")
    key_rek = "No Rek"
    key_nama = "Nama"
    
    # Slider untuk menentukan batas kemiripan nama khusus untuk No Rek yang sama
    similarity_threshold = st.slider(
        "Batas Minimal Kemiripan Nama (%):", 
        min_value=0, 
        max_value=100, 
        value=90, 
        step=5,
        help="Jika kemiripan nama untuk No Rek yang sama berada di bawah angka ini, nama akan dipisah ke kolom baru di kanan."
    )
    
    join_method = st.radio(
        "Metode Gabung No Rek:",
        options=["Tampilkan semua No Rek (Outer Join)", "Hanya No Rek yang ada di semua sheet (Inner Join)"],
        index=0
    )
    how_strategy = "outer" if "all" in join_method or "all" in join_method.lower() or "semua" in join_method.lower() else "inner"

with col2:
    st.subheader("📁 Unggah File")
    uploaded_files = st.file_uploader("Unggah File Excel Anda (.xlsx)", accept_multiple_files=True, type=["xlsx"])

# Fungsi internal untuk menghitung persentase kemiripan kata/nama
def hitung_kemiripan(str1, str2):
    if pd.isna(str1) or pd.isna(str2):
        return 0.0
    return SequenceMatcher(None, str(str1).strip().lower(), str(str2).strip().lower()).ratio() * 100

# 2. Proses Data
if uploaded_files:
    all_dfs = []
    
    # Ekstrak seluruh sheet dari file yang diupload
    for uploaded_file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                df.columns = df.columns.str.strip() # Bersihkan nama kolom dari spasi liar
                
                if key_rek in df.columns and key_nama in df.columns:
                    df[key_rek] = df[key_rek].astype(str).str.strip()
                    df[key_nama] = df[key_nama].astype(str).str.strip()
                    
                    # Simpan nama sheet asal di atribut internal dataframe
                    df.attrs['sheet_name'] = sheet_name
                    all_dfs.append(df)
                else:
                    st.warning(f"⚠️ Sheet '{sheet_name}' dilewati karena tidak memiliki kolom '{key_rek}' atau '{key_nama}'.")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    # 3. Logika Inti: Merge & Evaluasi Kemiripan per No Rek
    if all_dfs:
        # Gunakan sheet pertama sebagai baseline utama
        base_df = all_dfs[0].copy()
        base_sheet = all_dfs[0].attrs['sheet_name']
        
        # Looping untuk setiap sheet berikutnya yang akan ditempelkan ke samping
        for next_df in all_dfs[1:]:
            sheet_name = next_df.attrs['sheet_name']
            
            # Gabungkan sementara berdasarkan No Rek untuk membandingkan nama pada No Rek yang sama
            temp_merged = pd.merge(
                base_df[[key_rek, key_nama]], 
                next_df[[key_rek, key_nama]], 
                on=key_rek, 
                how='inner', 
                suffixes=('_base', '_next')
            )
            
            # Hitung skor kemiripan nama baris demi baris pada No Rek yang sama
            temp_merged['skor'] = temp_merged.apply(
                lambda r: hitung_kemiripan(r[f'{key_nama}_base'], r[f'{key_nama}_next']), axis=1
            )
            
            # Temukan daftar No Rek yang nama pasangannya di bawah standar (Gagal mirip)
            rek_berbeda_nama = temp_merged[temp_merged['skor'] < similarity_threshold][key_rek].tolist()
            
            # Buat salinan dataframe untuk dimodifikasi sebelum digabungkan resmi
            next_df_prepared = next_df.copy()
            
            # Buat kolom baru untuk menampung nama yang berbeda ke sebelah kanan
            kolom_nama_baru = f"{key_nama}_{sheet_name}"
            next_df_prepared[kolom_nama_baru] = None
            
            # Cari baris yang No Rek-nya bermasalah, pindahkan namanya ke kolom baru, lalu kosongkan kolom 'Nama' utamanya
            mask_beda = next_df_prepared[key_rek].isin(rek_berbeda_nama)
            next_df_prepared.loc[mask_beda, kolom_nama_baru] = next_df_prepared.loc[mask_beda, key_nama]
            next_df_prepared.loc[mask_beda, key_nama] = None
            
            # Beri suffix unik untuk kolom data lain (misal: Saldo, Alamat, dll) agar tidak bentrok antar sheet
            rename_dict = {}
            for col in next_df_prepared.columns:
                if col not in [key_rek, key_nama, kolom_nama_baru]:
                    rename_dict[col] = f"{col}_{sheet_name}"
            next_df_prepared = next_df_prepared.rename(columns=rename_dict)
            
            # Gabungkan secara resmi ke data induk utama
            base_df = pd.merge(base_df, next_df_prepared, on=key_rek, how=how_strategy)
            
            # Jika ada sisa pecahan kolom 'Nama' akibat join otomatis pandas (_x dan _y), rapikan kembali
            if f"{key_nama}_x" in base_df.columns and f"{key_nama}_y" in base_df.columns:
                base_df[key_nama] = base_df[f"{key_nama}_x"].fillna(base_df[f"{key_nama}_y"])
                base_df = base_df.drop(columns=[f"{key_nama}_x", f"{key_nama}_y"])

        # 4. Tampilkan Hasil Akhir di Layar Web
        st.success(f"🎉 Selesai memproses! Total hasil gabungan: {len(base_df)} baris.")
        
        with st.expander("👁️ Tampilkan Sampel Data Gabungan"):
            st.dataframe(base_df.head(10))
            
        # 5. Export ke format Excel siap pakai
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            base_df.to_excel(writer, index=False, sheet_name='Hasil_Gabungan_Nama')
        
        excel_bytes = buffer.getvalue()
        
        st.download_button(
            label="⬇️ Unduh File Excel Hasil Validasi",
            data=excel_bytes,
            file_name="merge_by_rek_and_similarity.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
