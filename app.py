import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher

st.set_page_config(page_title="Merge No Rek & Similarity Nama", layout="wide")

st.title("🗂️ Penggabung Excel (Validasi Nama per No Rek)")
st.write("Skrip ini mencocokkan **No Rek**. Jika baris No Rek yang sama memiliki **Nama** dengan kemiripan di bawah target slider, nama sheet pembanding akan dipecah ke kolom kanan baru. Kolom umum lainnya (Alamat, dll) akan otomatis ditimpa oleh sheet terbaru.")

# 1. Konfigurasi Kontrol di Sidebar / Halaman Utama
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("⚙️ Atur Batas Toleransi")
    key_rek = "No Rek"
    key_nama = "Nama"
    
    similarity_threshold = st.slider(
        "Batas Minimal Kemiripan Nama (%):", 
        min_value=0, max_value=100, value=90, step=5,
        help="Jika kemiripan nama di bawah angka ini, nama baru akan dipisah ke kolom kanan."
    )
    
    join_method = st.radio(
        "Metode Gabung No Rek:",
        options=["Tampilkan semua No Rek (Outer Join)", "Hanya No Rek yang ada di semua sheet (Inner Join)"],
        index=0
    )
    how_strategy = "outer" if "all" in join_method or "semua" in join_method.lower() else "inner"

with col2:
    st.subheader("📁 Unggah File")
    uploaded_files = st.file_uploader("Unggah File Excel Anda (.xlsx)", accept_multiple_files=True, type=["xlsx"])

def hitung_kemiripan(str1, str2):
    if pd.isna(str1) or pd.isna(str2):
        return 0.0
    return SequenceMatcher(None, str(str1).strip().lower(), str(str2).strip().lower()).ratio() * 100

# 2. Proses Data
if uploaded_files:
    all_dfs = []
    
    for uploaded_file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                df.columns = df.columns.str.strip()
                
                if key_rek in df.columns and key_nama in df.columns:
                    df[key_rek] = df[key_rek].astype(str).str.strip()
                    df[key_nama] = df[key_nama].astype(str).str.strip()
                    
                    df.attrs['sheet_name'] = sheet_name
                    all_dfs.append(df)
                else:
                    st.warning(f"⚠️ Sheet '{sheet_name}' dilewati karena tidak memiliki kolom '{key_rek}' atau '{key_nama}'.")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    # 3. Logika Inti: Merge & Evaluasi Kemiripan
    if all_dfs:
        base_df = all_dfs[0].copy()
        
        for next_df in all_dfs[1:]:
            sheet_name = next_df.attrs['sheet_name']
            
            temp_merged = pd.merge(
                base_df[[key_rek, key_nama]], next_df[[key_rek, key_nama]], 
                on=key_rek, how='inner', suffixes=('_base', '_next')
            )
            
            temp_merged['skor'] = temp_merged.apply(
                lambda r: hitung_kemiripan(r[f'{key_nama}_base'], r[f'{key_nama}_next']), axis=1
            )
            
            rek_berbeda_nama = temp_merged[temp_merged['skor'] < similarity_threshold][key_rek].tolist()
            
            next_df_prepared = next_df.copy()
            kolom_nama_baru = f"{key_nama}_{sheet_name}"
            next_df_prepared[kolom_nama_baru] = None
            
            # Pindahkan nama yang tidak lolos toleransi ke kolom pecahan, kosongkan kolom utama
            mask_beda = next_df_prepared[key_rek].isin(rek_berbeda_nama)
            next_df_prepared.loc[mask_beda, kolom_nama_baru] = next_df_prepared.loc[mask_beda, key_nama]
            next_df_prepared.loc[mask_beda, key_nama] = None
            
            # GABUNGKAN
            base_df = pd.merge(base_df, next_df_prepared, on=key_rek, how=how_strategy)
            
            # RESOLUSI OTOMATIS: Timpa semua kolom berduplikasi (_x vs _y)
            cols_x = [c for c in base_df.columns if c.endswith("_x")]
            
            for col_x in cols_x:
                nama_asli = col_x[:-2]
                col_y = f"{nama_asli}_y"
                
                if col_y in base_df.columns:
                    base_df[col_x] = base_df[col_y].fillna(base_df[col_x])
                    base_df = base_df.drop(columns=[col_y])
                    base_df = base_df.rename(columns={col_x: nama_asli})

        # 4. Output Web
        st.success(f"🎉 Selesai memproses! Total hasil gabungan: {len(base_df)} baris.")
        
        with st.expander("👁️ Tampilkan Sampel Data Gabungan"):
            st.dataframe(base_df.head(10))
            
        # 5. Export
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            base_df.to_excel(writer, index=False, sheet_name='Hasil_Gabungan')
        
        st.download_button(
            label="⬇️ Unduh File Excel Hasil Validasi",
            data=buffer.getvalue(),
            file_name="merge_by_rek_and_similarity.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
