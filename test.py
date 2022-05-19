import os
import pandas as pd
import xlrd
cols = ['Distributor Number', 'Reseller Name', 'Part Number', 'Description', 'Open Qty',
                    'Order Date', 'Required Delivery Date']
path_list = ["/Users/cameronlooney/Downloads/220511_APP_ORD.txt","/Users/cameronlooney/Documents/AOU Merge/AOU Merge Test Passed/VEN184_Apple_backlog_2022_05_11 7.XLS"]
#path  = "/Users/cameronlooney/Documents/name-test.xlsx"

df_list = []
for path in path_list:
    filename, file_extension = os.path.splitext(path)
    def read_to_df(path):
        if file_extension.lower() == ".xlsx":
            try:

                data_frame = pd.read_excel(path, sheet_name=0, engine="openpyxl")
                return data_frame
            except:
                "ERROR with .xlsx file"
        elif file_extension.lower() == ".xls":
            try:
                book = xlrd.open_workbook(filename=path)
                data_frame = pd.read_excel(book)
                return data_frame
            except:
                pass
            try:
                data_frame = pd.read_csv(path, sep='\t', encoding='latin1')
                return data_frame
            except:
                pass

        elif file_extension.lower() == ".txt":
            data_frame = pd.read_csv(path, sep='\t', encoding='latin1')
            return data_frame
        elif file_extension.lower() == ".csv":
            data_frame = pd.read_csv(path, sep=';', encoding='latin1')
            data_frame = data_frame.loc[:, ~data_frame.columns.str.contains('^Unnamed')]
            return data_frame


    data_frame = read_to_df(path)

    print(data_frame)

    part_num_names = ["Manufacturer Part Number", "PPN", "Part"]  # done
    distributor_num = ["Sold To", "Distributor Number"]  # done
    reseller_names = ["Reseller", "Name"]  # done
    required_date_names = ["ReqDelDate", "Order Ending Date"]
    open_qty_names = ["Open Quantity", "Qty open (SO)"]  # done
    description_names = ["Product Description"]
    order_date_names = ["Created on"]


    def fix_column_names(data_frame):
        x = list(data_frame)
        for i in x:
            if i in part_num_names:
                data_frame.rename(columns={i: 'Part Number'}, inplace=True)
            elif "part" in i.lower().strip():
                data_frame.rename(columns={i: 'Part Number'}, inplace=True)
            elif i.strip() == 'Part Number':
                data_frame.rename(columns={i: 'Part Number'}, inplace=True)

            # distributor_number fix
            if i in distributor_num:
                data_frame.rename(columns={i: 'Distributor Number'}, inplace=True)
            elif i.strip() == 'Distributor Number':
                data_frame.rename(columns={i: 'Distributor Number'}, inplace=True)
            # reseller_names fix
            if i in reseller_names:
                data_frame.rename(columns={i: 'Reseller Name'}, inplace=True)
            elif "reseller" in i.lower().strip():
                data_frame.rename(columns={i: 'Reseller Name'}, inplace=True)
            elif i.strip() == 'Reseller Name':
                data_frame.rename(columns={i: 'Reseller Name'}, inplace=True)

            # order date
            if i in order_date_names:
                data_frame.rename(columns={i: 'Order Date'}, inplace=True)
            elif i.strip() == 'Order Date':
                data_frame.rename(columns={i: 'Order Date'}, inplace=True)
            # open qty name fix

            if i in open_qty_names:
                data_frame.rename(columns={i: 'Open Qty'}, inplace=True)
            quantity_strings = ["open", "qty", "quantity"]
            if any(x in quantity_strings for x in i.lower().strip()):
                data_frame.rename(columns={i: 'Open Qty'}, inplace=True)
            elif i.strip() == 'Open Qty':
                data_frame.rename(columns={i: 'Open Qty'}, inplace=True)

            # delivery date fix
            if i in required_date_names:
                data_frame.rename(columns={i: 'Required Delivery Date'}, inplace=True)
            elif i.strip() == 'Required Delivery Date':
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
    print(path)
    print(list(data))
    def drop_extra_columns(data_frame):
        try:
            data_frame = data_frame[cols]
            return data_frame
        except:
            print("Missing Column")

    finished = drop_extra_columns(data)
    df_list.append(finished)
excl_merged = pd.concat(df_list, ignore_index=True)


