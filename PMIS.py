import streamlit as st
import pandas as pd
from datetime import datetime

def process_data(file1, file2, file3):
    """
    Processes the uploaded Excel files based on the original script's logic.
    """
    # Load dataframes
    excel_data1 = pd.ExcelFile(file1)
    df1 = excel_data1.parse('Work Order List')

    excel_data2 = pd.ExcelFile(file2)
    df2 = excel_data2.parse('Sheet1')
    df4 = excel_data2.parse('Function') # This df4 is used directly later

    excel_data3 = pd.ExcelFile(file3)
    df3 = excel_data3.parse('Ibase Data WO Open 1')

    # --- Data Processing (as per your original script) ---

    # Mendeteksi duplikat berdasarkan kolom 'product' dan menghapus baris yang memiliki status 'canceled' jika ada duplikat
    mask = df1.duplicated(subset=['Customer WOID'], keep=False) & (df1['Installation Status'] == 'Cancelled')
    df1 = df1[~mask]

    # Mengambil nilai kolom dari DataFrame
    values_df1 = set(df1.iloc[:, 5])
    values_df2 = set(df2.iloc[:, 22])

    # Mencari nilai yang berbeda di df1 dan df2
    different_values_in_df1 = values_df1 - values_df2
    df1 = df1[df1.iloc[:, 5].isin(different_values_in_df1)]

    # Menghapus seluruh huruf dari kolom 'PPPoE Username', hanya menyisakan angka
    df1['PPPoE Username'] = df1['PPPoE Username'].str.replace(r'\D', '', regex=True)

    # Mengganti nilai kota
    df1['Residence'] = df1['Residence'].replace({
        'GORONTALO': 'KAB. GORONTALO',
        'BENGKULU': 'KOTA BENGKULU',
        'LAHAT': 'KAB. LAHAT',
        'SIMALUNGUN': 'KAB. SIMALUNGUN',
        'SITUBONDO': 'KAB. SITUBONDO',
        'WONOSOBO': 'KAB. WONOSOBO',
        'JAMBI': 'KOTA JAMBI'
    })

    # Dictionary untuk mapping city ke region
    city_to_region = {
        'KOTA JAMBI': 'WEST',
        'KOTA BENGKULU': 'WEST',
        'KAB. SIMALUNGUN': 'WEST',
        'KAB. LAHAT': 'WEST',
        'KAB. WONOSOBO': 'CENTRAL',
        'KAB. SITUBONDO': 'EAST',
        'KAB. GORONTALO': 'SULAWESI'
    }

    # Tambahkan atau ganti nilai di kolom 'Region' berdasarkan nilai di kolom 'City'
    df1['Region'] = df1['Residence'].map(city_to_region)

    # Mengubah seluruh isi kolom
    df1['Vendor'] = 'TBG'
    df1['Status'] = 'Scheduled'

    # Membuat kolom baru dengan value yang sama
    df1['Product Type'] = 'XL SATU'
    df1['Area Type'] = 'Partnership - TBG'

    # Membuat duplikat kolom
    df1['Subcon'] = df1['Vendor']
    df1['City (Simplified)'] = df1['Residence']
    df1['Product Description'] = df1['Product ID']

    # Ubah nama kolom di df_wol agar sesuai dengan df_homeconnect
    df1_renamed = df1.rename(columns={
        'PPPoE Username': 'Customer iD',
        'WOID': 'WO Partner',
        'Residence' : 'City',
        'Cluster Name' : 'Cluster',
        'Customer WOID': 'WOID',
        'Homepass ID' : 'HP ID',
        'Product ID' : 'Charging Name',
        'Subscriber Name': 'Customer Name',
        'Status': 'STATUS WO',
        'Installation Date' : 'Latest Plan',
        'WO Created Date' : 'Input Data Date',
        'Reason' : 'WO Reason (only for Return)',
        'Time Slot' : 'SLOT TIME'
    })

    # Hanya pilih kolom di df2 yang ada di df1
    df1_filtered = df1_renamed[df2.columns.intersection(df1_renamed.columns)]

    # Gabungkan df1 dan df2_filtered
    df_combined = pd.concat([df2, df1_filtered], ignore_index=True, sort=False)

    # Ubah nama kolom di df3 agar sesuai
    df3 = df3.rename(columns={
        'SIM Card 1': 'SIM Card1',
        'SIM Card 2': 'SIM Card2',
    })

    # Tambahkan kolom baru 'Material ONT' dan 'Material STB' dengan kondisi sesuai
    df3['Material ONT'] = df3['SN ONT'].apply(lambda x: 'New' if pd.notna(x) else 'No ONT')
    df3['Material STB'] = df3['SN ONT'].apply(lambda x: 'No STB' if pd.notna(x) else 'No STB')

    # Tentukan kolom yang ingin digantikan
    kolom_diganti = [
        'SIM Card1', 'SIM Card2', 'Instalaton Remarks', 'SN ONT', 'STATUS WO',
        'Sub Status WO', 'WO Reason (only for Return)',
        'Active Date', 'Active Month', 'Active Years', 'Return / Cancel Date'
    ]

    # Menggabungkan DataFrame dengan menggantikan data di kolom tertentu dan menambah baris baru
    df_combined.set_index('WOID', inplace=True)
    df3.set_index('WOID', inplace=True)

    # Update hanya kolom yang ditentukan
    df_combined.update(df3[kolom_diganti])

    # Menambahkan baris baru dari df3 yang WOID-nya tidak ada di df_combined
    # Need to reset index on df_combined if it's already set to prevent issues when adding rows
    df_combined.reset_index(inplace=True)
    df_combined = pd.concat([df_combined, df3.loc[~df3.index.isin(df_combined['WOID']), kolom_diganti].reset_index()])


    # Reset index agar WOID kembali menjadi kolom biasa
    # Already reset above for concat, ensure no double reset if that causes issues.
    # df_combined.reset_index(inplace=True) # Potentially redundant if already done above for concat

    # Mengupdate kolom 'Sub Status WO' berdasarkan kondisi 'STATUS WO'
    df_combined.loc[df_combined['STATUS WO'] == 'Scheduled', 'Sub Status WO'] = 'Fix Scheduled'
    df_combined.loc[df_combined['STATUS WO'] == 'Installed', 'Material ONT'] = 'New'
    df_combined.loc[df_combined['STATUS WO'] == 'Installed', 'Material STB'] = 'No STB'


    # Mengubah kolom 'SIM Card' menjadi hanya angka
    df_combined['SIM Card1'] = df_combined['SIM Card1'].apply(lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notnull(x) else '')
    df_combined['SIM Card2'] = df_combined['SIM Card2'].apply(lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notnull(x) else '')

    # List of columns to convert as per user's instruction
    date_columns = [
        'Latest Plan',
        'Input Data Date',
        'Active Date',
        'Return / Cancel Date' # 'Latest Plan' listed twice, removed the duplicate here
    ]

    # Convert these columns to datetime and then format to 'dd/mmm/yyyy'
    for col in date_columns:
        df_combined[col] = pd.to_datetime(df_combined[col], errors='coerce').dt.strftime('%d/%b/%Y')

    # Menambahkan 'Completed' ke kolom 'Installation Remarks'
    df_combined.loc[df_combined['STATUS WO'] == 'Installed', 'Instalaton Remarks'] = df_combined['Active Date'] + " Completed"
    df_combined.loc[df_combined['STATUS WO'] == 'Scheduled', 'Instalaton Remarks'] = df_combined['Latest Plan'] + " On Plan"
    df_combined.loc[df_combined['STATUS WO'] == 'Return', 'Instalaton Remarks'] = df_combined['Return / Cancel Date'] + " " + df_combined['WO Reason (only for Return)']
    df_combined.loc[df_combined['STATUS WO'] == 'Hold', 'Instalaton Remarks'] = df_combined['Latest Plan'] + " Network Problem"


    # Menghapus isi kolom 'WO Reason (only for Return)' jika 'STATUS WO' adalah 'Scheduled'
    df_combined.loc[df_combined['STATUS WO'] == 'Scheduled', 'WO Reason (only for Return)'] = ''

    # Dapatkan urutan semua kolom
    cols = list(df_combined.columns)

    # Menukar posisi kolom 'Umur' dan 'Pekerjaan' - assuming this refers to 'WOID' and 'Customer iD' based on your code
    col1_name, col2_name = 'WOID', 'Customer iD'
    if col1_name in cols and col2_name in cols:
        idx1, idx2 = cols.index(col1_name), cols.index(col2_name)
        # Tukar posisi kolom
        cols[idx1], cols[idx2] = cols[idx2], cols[idx1]
    else:
        st.warning(f"Columns '{col1_name}' or '{col2_name}' not found for swapping.")


    # Atur ulang DataFrame berdasarkan urutan kolom yang baru
    df_final = df_combined[cols]

    return df_final, df4

# --- Streamlit UI ---
st.set_page_config(layout="wide", page_title="PMIS Data Processing")

st.title("🏡 PMIS Data Processing")
st.markdown("Upload your Excel files to process and generate the `HOMECONNECT.xlsx` output.")

st.header("Upload Your Files")
file_path1_upload = st.file_uploader("Upload Work Order List", type=["xlsx"])
file_path2_upload = st.file_uploader("Upload HomeConnect WO TBG", type=["xlsx"])
file_path3_upload = st.file_uploader("Upload Daily Report", type=["xlsx"])

if file_path1_upload and file_path2_upload and file_path3_upload:
    st.success("All files uploaded successfully! 🎉")

    if st.button("🚀 Process Data"):
        with st.spinner("Processing data... This might take a moment. ⏳"):
            try:
                df_final_output, df4_output = process_data(file_path1_upload, file_path2_upload, file_path3_upload)
                st.success("Data processing complete! ✅")

                st.subheader("Processed Data (Sheet1)")
                st.dataframe(df_final_output)

                # Generate download button
                output_filename = "HOMECONNECT.xlsx"
                
                # To save multiple sheets to an Excel file in memory
                output = pd.ExcelWriter(output_filename, engine='xlsxwriter')
                df_final_output.to_excel(output, sheet_name='Sheet1', index=False)
                df4_output.to_excel(output, sheet_name='Function', index=False, header=False)
                output.close()

                # Get the bytes from the Excel writer object
                with open(output_filename, "rb") as f:
                    excel_data_bytes = f.read()

                st.download_button(
                    label="⬇️ Download Processed Excel File",
                    data=excel_data_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.info(f"Your processed file '{output_filename}' is ready for download.")

            except Exception as e:
                st.error(f"An error occurred during processing: {e} 💔")
                st.exception(e)
else:
    st.info("Please upload all three Excel files to start the processing. ⬆️")

st.markdown("---")