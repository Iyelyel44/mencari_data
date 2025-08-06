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
        st.dataframe(df)

        # --- Bagian Pencarian Data ---
        st.write("---")
        st.subheader("Pilihan Fungsi Excel")
        
        # Tambahkan pilihan untuk user
        fungsi_pilihan = st.radio(
            "Pilih rumus yang ingin digunakan:",
            ('INDEX MATCH', 'COUNTIF')
        )
        
        st.subheader(f"Cari Data dengan {fungsi_pilihan}")

        # Memilih kolom untuk pencarian
        search_column = st.selectbox("Pilih kolom untuk pencarian", df.columns)
        
        # Logika untuk menampilkan kolom hasil hanya jika INDEX MATCH dipilih
        if fungsi_pilihan == 'INDEX MATCH':
            result_column = st.selectbox("Pilih kolom yang ingin ditampilkan hasilnya", df.columns)
        
        search_query = st.text_input(f"Masukkan data yang ingin dicari di kolom '{search_column}'")

        if st.button("Jalankan"):
            if search_query:
                # --- Logika Berdasarkan Pilihan User ---
                
                if fungsi_pilihan == 'INDEX MATCH':
                    # Logika untuk INDEX MATCH
                    hasil_pencarian = df[df[search_column].astype(str).str.contains(search_query, case=False, na=False)]
                    
                    if not hasil_pencarian.empty:
                        st.success("Data ditemukan!")
                        st.dataframe(hasil_pencarian[[search_column, result_column]])
                    else:
                        st.warning("Data tidak ditemukan.")
                
                elif fungsi_pilihan == 'COUNTIF':
                    # Logika untuk COUNTIF
                    # Hitung jumlah kemunculan
                    jumlah_data = len(df[df[search_column].astype(str).str.contains(search_query, case=False, na=False)])
                    
                    if jumlah_data > 0:
                        st.success(f"Data '{search_query}' ditemukan sebanyak {jumlah_data} kali.")
                    else:
                        st.warning(f"Data '{search_query}' tidak ditemukan.")

            else:
                st.warning("Mohon masukkan data yang ingin dicari.")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")

