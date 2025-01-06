import pandas as pd
import streamlit as st

st.title("Kertas Kerja Anggota Keluar")
st.write("""File ini berisikan TAK.xlsx dan TLP.xsx yang sudah di olah dengan OPTIMA serta sudah disatukan dengan N/A nya.""")

# Upload file
uploaded_tak = st.file_uploader("Upload TAK.xlsx", type="xlsx")
uploaded_tlp = st.file_uploader("Upload TLP.xlsx", type="xlsx")

if uploaded_tak and uploaded_tlp:
    try:
        # Load data
        tak_df = pd.read_excel(uploaded_tak)
        tlp_df = pd.read_excel(uploaded_tlp)

        # Proses Data
        combined_df = tak_df.copy()
        combined_df["No."] = range(1, len(tak_df) + 1)
        combined_df["Client ID"] = combined_df["ID ANGGOTA"]
        combined_df["Client Name"] = combined_df["NAMA"]
        combined_df["Center ID"] = combined_df["CENTER"]
        combined_df["Group ID"] = combined_df["KEL"]
        combined_df["Officer Name"] = combined_df["SL"]
        combined_df["Tanggal Keluar"] = combined_df["TRANS. DATE"]
        
        # VLOOKUP Total Pinjaman dari TLP.xlsx
        tlp_df = tlp_df[["ID ANGGOTA", "Db Total2"]]
        tlp_df = tlp_df.rename(columns={"ID ANGGOTA": "ID ANGGOTA TLP", "Db Total2": "Total Pinjaman"})
        combined_df = combined_df.merge(tlp_df, left_on="ID ANGGOTA", right_on="ID ANGGOTA TLP", how="left")
        combined_df["Total Pinjaman"] = combined_df["Total Pinjaman"].fillna(0)
        
        # Total Simpanan dari Cr Total
        combined_df["Total Simpanan"] = combined_df["Cr Total"]

        # Terima/Bayar = Total Pinjaman - Total Simpanan
        combined_df["Terima/ Bayar"] = combined_df["Total Simpanan"] - combined_df["Total Pinjaman"]

        # Kolom tambahan
        combined_df["Form AK"] = ""
        combined_df["Sesuai/ Tidak Sesuai"] = ""
        combined_df["Keterangan"] = ""

        # Kolom yang ingin disimpan
        final_df = combined_df[[
            "No.", "Client ID", "Client Name", "Center ID", "Group ID", 
            "Officer Name", "Tanggal Keluar", "Total Pinjaman", 
            "Total Simpanan", "Terima/ Bayar", "Form AK", 
            "Sesuai/ Tidak Sesuai", "Keterangan"
        ]]

        # Tombol unduh
        st.success("Data berhasil diproses!")
        st.dataframe(final_df)

        @st.cache_data
        def convert_df_to_excel(df):
            # Convert DataFrame ke Excel
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")
            processed_data = output.getvalue()
            return processed_data

        excel_data = convert_df_to_excel(final_df)
        st.download_button(
            label="Unduh KK Anggota Keluar",
            data=excel_data,
            file_name="KK Anggota Keluar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
else:
    st.warning("Silakan unggah kedua file untuk melanjutkan.")
