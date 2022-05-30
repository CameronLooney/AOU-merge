# import packages
from zipfile import ZipFile
import pandas as pd
import streamlit as st
import time
import os
import xlrd

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
cols = ['Distributor Number', 'Reseller Name', 'Part Number', 'Description', 'Open Qty',
        'Order Date', 'Required Delivery Date']
# if a file has been uploaded
file_not_read = []
file_not_fixed = []
if file_name is not None:
    # if the sidebar button is clicked
    if st.sidebar.button("Merge Files"):
        start = time.time()


        # unzip the uploaded files and get a list of the file paths
        def unzip():

            archive = ZipFile(file_name, 'r')
            with ZipFile(file_name, 'r') as zip:
                # printing all the contents of the zip file
                list_of_files = zip.namelist()
                return archive, list_of_files


        archive, list_of_files = unzip()


        # when zipping on mac os it generates MACOSX files which need to be removed
        def remove_MACOSX_files(list_of_files):
            # if path contains MACOSX drop it
            correct_list = [x for x in list_of_files if "MACOSX" not in x]
            correct_list = [x for x in correct_list if ".DS_Store" not in x]
            correct_list = [x for x in correct_list if "." in x]
            return correct_list


        path_list = remove_MACOSX_files(list_of_files)
        df_list = []
        file_not_read = []
        files_failed = []
        for i in path_list:

            def read_to_df(path):
                head, tail = os.path.split(path)

                try:
                    # .xlsx

                    xl = archive.open(path)
                    data_frame = pd.read_excel(xl, sheet_name=0, engine="openpyxl")

                    # this is to catch ingrammicro template (has annotations in line 1)

                    return data_frame
                except:
                    pass
                try:
                    xl = archive.open(path)
                    data_frame = pd.read_excel(xl)
                    return data_frame
                except:
                    pass

                try:

                    book = xlrd.open_workbook(filename=path)
                    st.write(book)
                    data_frame = pd.read_excel(book)
                    return data_frame
                except:
                    pass
                # txt/ csv
                try:
                    xl = archive.open(path)
                    data_frame = pd.read_csv(xl, sep='\t', encoding='latin1')

                    if len(data_frame.columns) <= 6:

                        pass
                    else:

                        return data_frame

                except:
                    pass
                try:
                    xl = archive.open(path)
                    data_frame = pd.read_csv(xl, sep=';', encoding='latin1')

                    data_frame = data_frame.loc[:, ~data_frame.columns.str.contains('^Unnamed')]
                    if "Order Ending Date" and "Customer requested delivery date" in data_frame:
                        data_frame = data_frame.drop('Order Ending Date', 1)
                    return data_frame
                except:
                    file_not_read.append(path)
                    return file_not_read


            data_frame = read_to_df(i)

            part_num_names = ["Manufacturer Part Number", "PPN", "Part", "Apple Part number"]  # done
            distributor_num = ["Sold To", "Distributor Number", "Distributor Sold to #"]  # done
            reseller_names = ["Reseller", "Name"]  # done
            required_date_names = ["ReqDelDate", "Order Ending Date", "End customer required delivery date",
                                   "Required delivery date", "Customer requested delivery date"]
            open_qty_names = ["Open Quantity", "Qty open (SO)", "Open Order Qty"]  # done
            description_names = ["Product Description"]
            order_date_names = ["Created on", "Date of order placement", "order"]

            try:

                data_frame = data_frame.dropna(axis=1, how='all')
                if any("Unnamed" in x for x in list(data_frame)):
                    data_frame.rename(columns=data_frame.iloc[0], inplace=True)
                    data_frame.drop([0], inplace=True)


            except:
                pass


            def fix_column_names(data_frame):
                x = list(data_frame)
                for i in x:
                    if i in part_num_names:
                        data_frame.rename(columns={i: 'Part Number'}, inplace=True)
                    elif "part" in i.lower().strip():
                        data_frame.rename(columns={i: 'Part Number'}, inplace=True)
                    elif i.strip().lower() == 'part number':
                        data_frame.rename(columns={i: 'Part Number'}, inplace=True)

                    # distributor_number fix
                    if i in distributor_num:
                        data_frame.rename(columns={i: 'Distributor Number'}, inplace=True)
                    elif i.strip().lower() == 'distributor number':
                        data_frame.rename(columns={i: 'Distributor Number'}, inplace=True)
                    # reseller_names fix
                    if i in reseller_names and "#" not in i:
                        data_frame.rename(columns={i: 'Reseller Name'}, inplace=True)
                    elif "reseller" in i.lower().strip()  and "#" not in i:
                        data_frame.rename(columns={i: 'Reseller Name'}, inplace=True)
                    elif i.strip().lower() == 'reseller name':
                        data_frame.rename(columns={i: 'Reseller Name'}, inplace=True)

                    # order date
                    if i in order_date_names:
                        data_frame.rename(columns={i: 'Order Date'}, inplace=True)
                    elif i.strip().lower() == 'order date':
                        data_frame.rename(columns={i: 'Order Date'}, inplace=True)
                    elif "Order Date" not in x:
                        if i == "Order Entry Date":
                            data_frame.rename(columns={i: 'Order Date'}, inplace=True)

                    # open qty name fix

                    if i in open_qty_names:
                        data_frame.rename(columns={i: 'Open Qty'}, inplace=True)
                    quantity_strings = ["open", "qty", "quantity"]
                    if any(x in quantity_strings for x in i.lower().strip()):
                        data_frame.rename(columns={i: 'Open Qty'}, inplace=True)
                    elif i.strip().lower() == 'open qty':
                        data_frame.rename(columns={i: 'Open Qty'}, inplace=True)

                    # delivery date fix
                    if i in required_date_names:
                        data_frame.rename(columns={i: 'Required Delivery Date'}, inplace=True)
                    elif i.strip().lower() == 'required delivery date':
                        data_frame.rename(columns={i: 'Required Delivery Date'}, inplace=True)

                    # description fix
                    if i in part_num_names:
                        data_frame.rename(columns={i: 'Description'}, inplace=True)
                    elif "description" in i.lower().strip():
                        data_frame.rename(columns={i: 'Description'}, inplace=True)
                    elif i.strip() == 'Description':
                        data_frame.rename(columns={i: 'Description'}, inplace=True)
                return data_frame


            data = fix_column_names(data_frame)


            def drop_extra_columns(data_frame):
                try:
                    data_frame = data_frame[cols]
                    st.write(i + "successfully completed")
                    return data_frame
                except:
                    return file_not_fixed.append(data_frame)


            finished = drop_extra_columns(data)

            try:
                if len(finished.columns) != 7:
                    file_not_fixed.append(finished)
                    files_failed.append(i)
                elif len(finished.columns) == 7:
                    df_list.append(finished)
            except:
                files_failed.append(i)

            # compare dataframe column headers to list and if element in list is missing from dataframe create an empty column in the same position


        try:
            excl_merged = pd.concat(df_list, ignore_index=True)
            excl_merged = excl_merged.drop_duplicates(keep='first')
            st.success("Merger was sucessfully completed")
        except:
            st.error("Error: The merge was not completed")

        if len(files_failed)>0:
            st.error("The following files were not read properly:")
            st.write(files_failed)



        def excel(excl_merged):
            import io
            buffer = io.BytesIO()
            writer = pd.ExcelWriter(buffer, date_format='yyyy-mm-dd', datetime_format='yyyy-mm-dd')

            excl_merged.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']

            # Get the dimensions of the dataframe.
            (max_row, max_col) = excl_merged.shape

            # Set the column widths, to make the dates clearer.
            worksheet.set_column(1, max_col, 20)
            writer.save()
            return buffer


        to_excel = excel(excl_merged)


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

