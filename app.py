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
file_counter = 0
merged_rows = 0
# if file is uploaded
if file_name is not None:
    # if the sidebar button is clicked
    if st.sidebar.button("Merge Files"):
        with st.spinner("Merging Files..."):

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
                # drop file if it contains MACOSX
                correct_list = [x for x in list_of_files if "MACOSX" not in x]
                # drop file if it contains .DS_store
                correct_list = [x for x in correct_list if ".DS_Store" not in x]
                # drop picture files
                correct_list = [x for x in correct_list if ".jpg" not in x]
                correct_list = [x for x in correct_list if ".png" not in x]
                # drop file if it doesnt contain a "." (not a proper file to format)
                correct_list = [x for x in correct_list if "." in x]
                return correct_list

            # get file paths of files that passed the tests
            path_list = remove_MACOSX_files(list_of_files)
            # initalise empty lists
            df_list = []
            file_not_read = []
            files_failed = []
            for i in path_list:

                # this is incredibly ugly but the files have completely different formats and extensions so
                # for now I have hard coded to catch as many files as possible
                # Changes to files can result in them not being formatted
                def read_to_df(path):
                    # check if the file is .xlsx, if so read to df , if not try not extension
                    try:
                        xl = archive.open(path)
                        data_frame = pd.read_excel(xl, sheet_name=0, engine="openpyxl")


                        return data_frame
                    except:
                        pass
                    # try open the file as .xls, there are two methods of opening .xls as some files were not captured if
                    # just one open attempt was used
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
                    # try open the file as .txt or .csv.
                    # Files have different delimiters so we will try several
                    # try tab delimiter, if the len is too short we know we didnt split it correctly and used wrong delimiter
                    try:
                        xl = archive.open(path)
                        data_frame = pd.read_csv(xl, sep='\t', encoding='latin1')

                        if len(data_frame.columns) <= 6:

                            pass
                        else:

                            return data_frame

                    except:
                        pass
                    # try semi colon delimiter, if the len is too short we know we didnt split it correctly and used wrong delimiter
                    try:
                        xl = archive.open(path)
                        data_frame = pd.read_csv(xl, sep=';', encoding='latin1')
                        # this is a hard coded exception as one file has multiple trailing semi colons that produce empty columns
                        data_frame = data_frame.loc[:, ~data_frame.columns.str.contains('^Unnamed')]
                        # hard coded exception as one file has too many date columns and they are mislabelled
                        if "Order Ending Date" and "Customer requested delivery date" in data_frame:
                            data_frame = data_frame.drop('Order Ending Date', 1)
                        return data_frame
                    except:
                        # if we get here we have tried .xlsx, .csv . xls and .txt , if we couldnt read the file at this stage we add it to our
                        # list of files we couldnt process
                        file_not_read.append(path)
                        return file_not_read


                data_frame = read_to_df(i)

                # again, this is extremely ugly but seemed like the best solution
                # each file uses a variety of names so I went through and added as many variations to a list to compare against.
                # possibly name column based on content in future
                part_num_names = ["Manufacturer Part Number", "PPN", "Part", "Apple Part number"]  # done
                distributor_num = ["Sold To", "Distributor Number", "Distributor Sold to #"]  # done
                reseller_names = ["Reseller", "Name"]  # done
                required_date_names = ["ReqDelDate", "Order Ending Date", "End customer required delivery date",
                                       "Required delivery date", "Customer requested delivery date"]
                open_qty_names = ["Open Quantity", "Qty open (SO)", "Open Order Qty"]  # done
                description_names = ["Product Description"]
                order_date_names = ["Created on", "Date of order placement", "order"]

                try:
                    # several files start a line down or put random headings above the headings
                    # this is an exception to catch files that dont list all column headings in first row
                    data_frame = data_frame.dropna(axis=1, how='all')
                    if any("Unnamed" in x for x in list(data_frame)):
                        data_frame.rename(columns=data_frame.iloc[0], inplace=True)
                        data_frame.drop([0], inplace=True)


                except:
                    pass

                # here we do name checks and try fix column headers so we can join the dataframes successfully
                def fix_column_names(data_frame):
                    try:
                        # get list of headers
                        x = list(data_frame)
                        for i in x:
                            # check if header is in list of variations
                            # check if key word is in header
                            # try strip white spaces and capitalisation
                            # if we dont find a match compare against not possible heading
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
                    except:
                        # we couldnt match the column to any of the possible headings
                        df = pd.DataFrame(columns =["error"])
                        return df

                # apply name fixes to columns
                data = fix_column_names(data_frame)

                # many files contain redundant columns so they need to be dropped before we merge dfs
                def drop_extra_columns(data_frame):
                    try:
                        # try make new df with only necessary columns
                        # if this fails we are missing atleast one required column
                        data_frame = data_frame[cols]
                        return data_frame
                    except:
                        pass
                    try:
                        # introduced a one missing column limit. If the file is missing one column we will impute
                        # Not Provided. This might be changed

                        comparison = list(set(cols).intersection(list(data_frame)))
                        if len(comparison) == 6:
                            # loop through and find which of the columns is missing
                            col_list = list(data_frame)
                            for k in cols:
                                if k not in col_list:
                                    data_frame[k] = "Not Provided"
                            data_frame = data_frame[cols]
                            return data_frame
                    except:
                        return file_not_fixed.append(data_frame)



                finished = drop_extra_columns(data)
                # ensure the df has 7 columns as it should after being processed
                # if it doesnt we will add it to list of dfs not yet fixed
                try:
                    if len(finished.columns) != 7:
                        file_not_fixed.append(finished)
                        files_failed.append(i)
                    # if it has exaclty 7 then the file has been processed successfully
                    elif len(finished.columns) == 7:
                        df_list.append(finished)
                        file_counter+=1
                # exception if the file isnt a df, this means it couldnt be read
                except:
                    files_failed.append(i)


            try:
                # the files that have been processed successfully are merged
                # merge th
                excl_merged = pd.concat(df_list, ignore_index=True)
                excl_merged =  excl_merged.dropna(how='all')
                end = time.time()
                time_taken = (end - start)

                # overview of successfully process
                st.success("Merger was completed\n\n"
                           "Files merged: " + str(file_counter) +
                           "\n\nTotal rows merged: " + str(f"{len(excl_merged.index):,}") +
                           "\n\nTime taken: " + str(round(time_taken,2)) + " seconds")

            # something went wrong and we couldnt comlete the merging
            except:
                st.error("Error: The merge was not completed")
            # if we have atleast one file that wasnt merged print warning along with the file name so it can be merged manually
            if len(files_failed)>0:
                st.error("The following files were not read properly:")
                st.write(files_failed)


            # pretty and adjust file so it can be sent to excel file
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

            try:
                to_excel = excel(excl_merged)
            except:
                pass

            # download file
            def download(buffer):
                st.download_button(
                    label="Download Excel worksheets",
                    data=buffer,
                    file_name="Merged.xlsx",
                    mime="application/vnd.ms-excel"
                )

            try:
                download(to_excel)
            except:
                # pray we never reach here
                st.error("An unexpected error has occurred :( \n"
                         "Please reach out to cameron_j_looney@apple.com for technical support")

