import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher

st.set_page_config(page_title="Merge No Rek & Dynamic Similarity", layout="wide")

st.title("🗂️ Penggabung Excel (Validasi Dinamis per No Rek)")
st.write("Skrip ini menggabungkan data berbasis **No Rek** (Wajib 100% sama). Anda bisa mengunggah file Excel atau langsung **Copy-Paste tabel** dari layar Excel Anda.")

# 1. Konfigurasi Kontrol di Sidebar
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("⚙️ Atur Batas Toleransi")
    key_rek = "No Rek" 
    
    target_col = st.text_input(
        "Kolom yang Diuji Kemiripannya:", value="Nama",
        help="Ketik persis nama kolom yang ada di Excel Anda."
    )
    
    label_slider = f"Batas Kemiripan '{target_col}' (%):" if target_col else "Batas Kemiripan (%):"
    similarity_threshold = st.slider(
        label_slider, min_value=0, max_value=100, value=90, step=5,
        help=f"Jika kemiripan isi kolom '{target_col}' di bawah angka ini, datanya akan dipecah ke kolom baru."
    )
    
    join_method = st.radio(
        "Metode Gabung No Rek:",
        options=["Tampilkan semua No Rek (Outer Join)", "Hanya No Rek yang ada di semua sheet (Inner Join)"],
        index=0
    )
    how_strategy = "outer" if "all" in join_method or "semua" in join_method.lower() else "inner"

if not target_col.strip():
    st.warning("⚠️ Tentukan dahulu nama kolom yang ingin diuji pada kotak pengaturan di sebelah kiri.")
    st.stop()

def hitung_kemiripan(str1, str2):
    if pd.isna(str1) or pd.isna(str2):
        return 0.0
    return SequenceMatcher(None, str(str1).strip().lower(), str(str2).strip().lower()).ratio() * 100


# 2. Area Input Data (Pilih Mode)
with col2:
    st.subheader("📥 Masukkan Data")
    metode_input = st.radio(
        "Pilih Cara Input:",
        ["📁 Unggah File Excel (.xlsx)", "📋 Copy-Paste Langsung dari Excel"],
        horizontal=True
    )
    
    all_dfs = []

    # --- JALUR A: VIA UPLOAD EXCEL ---
    if metode_input == "📁 Unggah File Excel (.xlsx)":
        uploaded_files = st.file_uploader("Unggah File Excel Anda", accept_multiple_files=True, type=["xlsx"])
        if uploaded_files:
            for uploaded_file in uploaded_files:
                try:
                    excel_file = pd.ExcelFile(uploaded_file)
                    for sheet_name in excel_file.sheet_names:
                        df = pd.read_excel(excel_file, sheet_name=sheet_name)
                        df.columns = df.columns.str.strip()
                        
                        if key_rek in df.columns and target_col in df.columns:
                            df[key_rek] = df[key_rek].astype(str).str.strip()
                            df[target_col] = df[target_col].astype(str).str.strip()
                            df.attrs['sheet_name'] = sheet_name
                            all_dfs.append(df)
                        else:
                            st.warning(f"⚠️ Sheet '{sheet_name}' dilewati: tidak ada kolom '{key_rek}' atau '{target_col}'.")
                except Exception as e:
                    st.error(f"Gagal membaca file: {e}")

    # --- JALUR B: VIA COPY PASTE VISUAL ---
    else:
        jml_tabel = st.number_input("Berapa tabel yang ingin digabung?", min_value=2, max_value=5, value=2)
        tabs = st.tabs([f"Tabel {i+1}" for i in range(jml_tabel)])
        
        for i, tab in enumerate(tabs):
            with tab:
                st.caption("💡 **Cara pakai:** Buka Excel -> Block tabel beserta baris judul kolomnya -> Copy (Ctrl+C) -> Paste (Ctrl+V) di kotak bawah ini.")
                txt = st.text_area(f"Paste isi Tabel {i+1}:", height=130, key=f"paste_{i}", placeholder="No Rek\tNama\tAlamat\n100234\tBudi\tJakarta...")
                
                if txt.strip():
                    try:
                        # Auto-detect pemisah: Coba Tab (\t) dulu, kalau gagal coba koma/titik koma
                        df_p = pd.read_csv(io.StringIO(txt), sep='\t')
                        if len(df_p.columns) == 1 and ',' in txt:
                            df_p = pd.read_csv(io.StringIO(txt), sep=',')
                        elif len(df_p.columns) == 1 and ';' in txt:
                            df_p = pd.read_csv(io.StringIO(txt), sep=';')
                            
                        df_p.columns = df_p.columns.str.strip()
                        df_p.dropna(how='all', inplace=True) # Hapus baris kosong di bawah
                        
                        if key_rek in df_p.columns and target_col in df_p.columns:
                            df_p[key_rek] = df_p[key_rek].astype(str).str.strip()
                            df_p[target_col] = df_p[target_col].astype(str).str.strip()
                            
                            df_p.attrs['sheet_name'] = f"Tabel_Paste_{i+1}"
                            all_dfs.append(df_p)
                            
                            st.dataframe(df_p.head(4), use_container_width=True)
                        else:
                            st.error(f"⚠️ Teks yang dipaste tidak memiliki header kolom '{key_rek}' atau '{target_col}'.")
                    except Exception as err:
                        st.error(f"Gagal membaca format teks: {err}")


# 3. Logika Inti: Merge & Evaluasi
if len(all_dfs) >= 2:
    base_df = all_dfs[0].copy()
    
    for next_df in all_dfs[1:]:
        sheet_name = next_df.attrs['sheet_name']
        
        temp_merged = pd.merge(
            base_df[[key_rek, target_col]], next_df[[key_rek, target_col]], 
            on=key_rek, how='inner', suffixes=('_base', '_next')
        )
        
        temp_merged['skor'] = temp_merged.apply(
            lambda r: hitung_kemiripan(r[f'{target_col}_base'], r[f'{target_col}_next']), axis=1
        )
        
        rek_gagal_mirip = temp_merged[temp_merged['skor'] < similarity_threshold][key_rek].tolist()
        
        next_df_prepared = next_df.copy()
        kolom_pecahan_baru = f"{target_col}_{sheet_name}"
        next_df_prepared[kolom_pecahan_baru] = None
        
        mask_beda = next_df_prepared[key_rek].isin(rek_gagal_mirip)
        next_df_prepared.loc[mask_beda, kolom_pecahan_baru] = next_df_prepared.loc[mask_beda, target_col]
        next_df_prepared.loc[mask_beda, target_col] = None
        
        base_df = pd.merge(base_df, next_df_prepared, on=key_rek, how=how_strategy)
        
        cols_x = [c for c in base_df.columns if c.endswith("_x")]
        for col_x in cols_x:
            nama_asli = col_x[:-2]
            col_y = f"{nama_asli}_y"
            if col_y in base_df.columns:
                base_df[col_x] = base_df[col_y].fillna(base_df[col_x])
                base_df = base_df.drop(columns=[col_y])
                base_df = base_df.rename(columns={col_x: nama_asli})

    # 4. Output Web
    st.divider()
    st.success(f"🎉 Selesai memproses! Total hasil gabungan: {len(base_df)} baris.")
    
    with st.expander("👁️ Tampilkan Sampel Data Gabungan", expanded=True):
        st.dataframe(base_df, use_container_width=True)
        
    # 5. Export
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        base_df.to_excel(writer, index=False, sheet_name='Hasil_Gabungan')
    
    st.download_button(
        label="⬇️ Unduh File Excel Hasil Validasi",
        data=buffer.getvalue(),
        file_name="merge_dynamic_similarity.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif len(all_dfs) == 1:
    st.info("💡 Data pertama sudah masuk. Silakan isi minimal 1 tabel lagi untuk memicu penggabungan.")
