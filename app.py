import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gabung Data No Rek", layout="centered")

st.title("🔗 Penggabung Sheet Excel (Match Kolom 'No Rek')")
st.write("Skrip ini akan mencocokkan nomor rekening dan menggabungkan data ke arah kanan.")

# Input nama kolom kunci (default: No Rek)
key_column = st.text_input("Nama kolom yang menjadi kunci kecocokan:", value="No Rek")

# 1. Upload File
uploaded_files = st.file_uploader("Pilih file Excel (.xlsx)", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    dfs_to_merge = []
    
    # 2. Ekstrak semua sheet menjadi list DataFrame terpisah
    for file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(file)
            for sheet in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet)
                
                # Bersihkan spasi di nama kolom agar pencocokan akurat
                df.columns = df.columns.str.strip()
                
                if key_column in df.columns:
                    # Pastikan tipe data No Rek seragam (diubah ke string agar tidak terpotong nol di depan)
                    df[key_column] = df[key_column].astype(str).str.strip()
                    
                    # Beri akhiran unik pada kolom lain agar tahu asal sheet-nya
                    suffix_name = f"_{sheet}"
                    cols_to_rename = {col: f"{col}{suffix_name}" for col in df.columns if col != key_column}
                    df = df.rename(columns=cols_to_rename)
                    
                    dfs_to_merge.append((sheet, df))
                else:
                    st.warning(f"⚠️ Sheet '{sheet}' di file {file.name} dilewati karena tidak punya kolom '{key_column}'.")
        except Exception as e:
            st.error(f"Gagal membaca file {file.name}: {e}")

    # 3. Proses Penggabungan (Merge Horizontal)
    if dfs_to_merge:
        # Gunakan jenis merge 'outer' agar rekening yang hanya ada di salah satu sheet tetap terbawa
        # Pilih 'how=outer' untuk semua data, atau 'how=inner' jika hanya ingin data yang ada di SEMUA sheet.
        merge_type = st.radio("Metode Penggabungan Data:", 
                              options=["Semua data digabungkan (Outer Join)", "Hanya data yang cocok di semua sheet (Inner Join)"],
                              index=0)
        
        how_method = "outer" if "Outer" in merge_type else "inner"
        
        # Mulai gabungkan dari DataFrame pertama
        master_df = dfs_to_merge[0][1]
        
        for name, df in dfs_to_merge[1:]:
            master_df = pd.merge(master_df, df, on=key_column, how=how_method)
            
        st.success(f"🎉 Berhasil mencocokkan {len(dfs_to_merge)} sheet! Total baris hasil gabungan: {len(master_df)}")
        
        # Tampilkan Pratinjau
        with st.expander("👁️ Lihat Pratinjau Data Gabungan Ke Samping"):
            st.dataframe(master_df.head(10))
            
        # 4. Tombol Download Hasil
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            master_df.to_excel(writer, index=False, sheet_name='Hasil_Gabungan_Kanan')
        
        excel_data = buffer.getvalue()
        
        st.download_button(
            label="⬇️ Unduh File Gabungan Kanan",
            data=excel_data,
            file_name="data_gabungan_horizontal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
