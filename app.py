import streamlit as st
import pandas as pd

# Menyiapkan tampilan dasar halaman web
st.title("Pencari Data dari File Excel")
st.write("Unggah file Excel Anda, pilih sheet, dan cari data yang diinginkan.")

# --- Bagian Unggah File ---
uploaded_file = st.file_uploader("Pilih file Excel (format .xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Membaca semua sheet dari file Excel
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)
        
        # --- Bagian Pemilihan Sheet ---
        sheet_names = list(excel_data.keys())
        selected_sheet = st.selectbox("Pilih Sheet", sheet_names)
        
        # Mengambil DataFrame dari sheet yang dipilih
        df = excel_data[selected_sheet]
        
        st.write("---")
        st.subheader(f"Data dari sheet '{selected_sheet}':")
        st.dataframe(df.head()) # Menampilkan 5 baris pertama sebagai pratinjau

        # --- Bagian Pencarian Data ---
        st.write("---")
        st.subheader("Cari Data")
        
        # Memilih kolom untuk pencarian
        search_column = st.selectbox("Pilih kolom untuk pencarian", df.columns)
        
        # Memilih kolom yang ingin ditampilkan hasilnya
        result_column = st.selectbox("Pilih kolom yang ingin ditampilkan", df.columns)
        
        search_query = st.text_input(f"Masukkan data yang ingin dicari di kolom '{search_column}'")

        if st.button("Cari"):
            if search_query:
                # Logika pencarian yang meniru INDEX MATCH
                # Mencari baris yang cocok dengan kueri pencarian
                hasil_pencarian = df[df[search_column].astype(str).str.contains(search_query, case=False, na=False)]
                
                if not hasil_pencarian.empty:
                    st.success("Data ditemukan!")
                    st.dataframe(hasil_pencarian[[search_column, result_column]]) # Menampilkan hasil dalam tabel
                else:
                    st.warning("Data tidak ditemukan. Coba kata kunci lain.")
            else:
                st.warning("Mohon masukkan data yang ingin dicari.")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")