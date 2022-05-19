# import packages
from zipfile import ZipFile
import pandas as pd
import streamlit as st
import time
import os

# Streamlit Title
st.markdown('''
# **Excel File Merger**
---
''')
# upload zipfile
with st.sidebar.header('1. Upload your ZIP file'):
    file_name = st.sidebar.file_uploader("Excel-containing ZIP file", type=["zip"])
    st.sidebar.markdown("""

""")

if file_name is not None:
    if st.sidebar.button("Merge Files"):
        start = time.time()
        def unzip_and_merge():
            # columns to keep
            cols = ['Distributor Number', 'Reseller Name', 'Part Number', 'Description', 'Open Qty',
                    'Order Date', 'Required Delivery Date']
            part_num_names = ["Manufacturer Part Number", "PPN","Part"]
            # keep user updated
            st.spinner(text="In progress...")

            # list to append each dataframe to
            df_list = []
            #file types we will accept ( wont merge but wont get rid of like MACOSX files)
            file_types = [".xlsx", ".txt", ".xls"]
            # read in zipfiles
            archive = ZipFile(file_name, 'r')

            with ZipFile(file_name, 'r') as zip:
                # printing all the contents of the zip file
                # list of all excel files
                excel_list = zip.namelist()

                # all files processed successfully
                processed_list = []
                # files couldnt be processed
                error_list = []
                # counter
                files_merged = 0
                #excel_list = [ x for x in excel_list if ".xlsx" in x ]
                # if file type in name keep
                excel_list = [x for x in excel_list if any(word in x for word in file_types)]
                # drop macosx files
                excel_list = [x for x in excel_list if "MACOSX" not in x]


                for i in excel_list:
                    try:
                        # split path and name for lists
                        head, tail = os.path.split(i)
                        xl = archive.open(i)
                        # attempt to read as dataframe
                        data_frame = pd.read_excel(xl, sheet_name = 0,engine = "openpyxl")
                        # get list of the names and check against variations
                        # check the excel files contains all columns required

                        result = all(elem in list(data_frame) for elem in cols)
                        if result:
                            # if it has all columns add the df to list
                            df_list.append(data_frame)
                            processed_list.append(tail)
                            files_merged += 1
                        else:
                            # if its missing column add to error list
                            error_list.append(tail)

                    except:
                        head,tail = os.path.split(i)
                        error_list.append(tail)
                # join all the processed dataframes
                excl_merged = pd.concat(df_list, ignore_index=True)
            return excl_merged, processed_list, error_list, files_merged

        df ,processed, error, num_files_merged = unzip_and_merge()
        # Print each of the files that werent processed
        st.header("Files not Processed")
        for i in error:
            st.write(i)





        # drop cols incase there were extra
        def drop_columns(excl_merged):
                cols = ['Distributor Number', 'Reseller Name', 'Part Number', 'Description', 'Open Qty',
                        'Order Date'    ,'Required Delivery Date']
                excl_merged = excl_merged[cols]
                excl_merged['Order Date'] = pd.to_datetime(excl_merged['Order Date'], format='%Y-%m-%d')
                excl_merged['Required Delivery Date'] = pd.to_datetime(excl_merged['Required Delivery Date'])
                return excl_merged
        df_dropped = drop_columns(df)
        # write to excel
        def excel(excl_merged):
                import io
                buffer = io.BytesIO()
                writer = pd.ExcelWriter(buffer, date_format='yyyy-mm-dd',datetime_format='yyyy-mm-dd')

                excl_merged.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']

                # Get the dimensions of the dataframe.
                (max_row, max_col) = excl_merged.shape

                # Set the column widths, to make the dates clearer.
                worksheet.set_column(1, max_col, 20)
                writer.save()
                return buffer
        to_excel = excel(df_dropped)
        def download(buffer):
                st.download_button(
                    label="Download Excel worksheets",
                    data=buffer,
                    file_name="Merged.xlsx",
                    mime="application/vnd.ms-excel"
                )
        download(to_excel)

        end = time.time()
        time_taken = (end - start)

        def print_time():
            st.success("**Merge Complete!** \n\n Time Taken: "  + str(round(time_taken, 2)) + " seconds \n\n No. Files Merged: " + str(num_files_merged)
                       + "\n\nNo. Records: " + f"{len(df_dropped.index):,}")

        print_time()



