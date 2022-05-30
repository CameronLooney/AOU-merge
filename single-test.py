import os

import pandas as pd
import xlrd
cols = ['Distributor Number', 'Reseller Name', 'Part Number', 'Description', 'Open Qty',
                    'Order Date', 'Required Delivery Date']
path_list = ["/Users/cameronlooney/Downloads/220511_APP_ORD.txt","/Users/cameronlooney/Documents/AOU Merge/AOU Merge Test Passed/VEN184_Apple_backlog_2022_05_11 7.XLS"]
#path  = "/Users/cameronlooney/Documents/name-test.xlsx"

df_list = []

#path = "/Users/cameronlooney/Downloads/APPLE_B_A_2022-05-12.CSV"
path = "/Users/cameronlooney/Documents/AOU Merge/AOU Merge Test Passed/Apple Backlog All_34_1896914398520070056.xlsx"

filename, file_extension = os.path.splitext(path)
def column_check(data_frame):
    try:
        data_frame =data_frame[cols]
        return True
    except:
        return False
def read_to_df(path):

    try:
        # .xlsx

        data_frame = pd.read_excel(path, sheet_name=0, engine="openpyxl")


        return data_frame
    except:
        pass

    try:
        # .xls
        book = xlrd.open_workbook(filename=path)
        data_frame = pd.read_excel(book)
        return data_frame
    except:
        pass
    # txt/ csv
    try:
        data_frame = pd.read_csv(path, sep='\t', encoding='latin1')
        print("here")
        if len(data_frame.columns) <=6:
            pass
        else:
            return data_frame

    except:
        pass
    try:
        data_frame = pd.read_csv(path, sep=';', encoding='latin1')
        data_frame = data_frame.loc[:, ~data_frame.columns.str.contains('^Unnamed')]
        if "Order Ending Date" and "Customer requested delivery date" in data_frame:
            data_frame = data_frame.drop('Order Ending Date', 1)
        return data_frame
    except:
        return "Error"
data_frame = read_to_df(path)

part_num_names = ["Manufacturer Part Number", "PPN","Part","Apple Part number"]# done
distributor_num = ["Sold To","Distributor Number","Distributor Sold to #"]#done
reseller_names = ["Reseller","Name"]#done
required_date_names = ["ReqDelDate","Order Ending Date","End customer required delivery date","Required delivery date","Customer requested delivery date"]
open_qty_names = ["Open Quantity","Qty open (SO)","Open Order Qty"]#done
description_names = ["Product Description"]
order_date_names = ["Created on","Date of order placement","order"]



def remove_nans(df):
    x = df.dropna(how = 'all', axis = 1)
    x = x.dropna(how = 'all', axis = 0)
    x = x.reset_index(drop = True)
    return x

'''
headers = data_frame.iloc[0]
print(headers)
new_df  = pd.DataFrame(data_frame.values[1:], columns=headers)
print(new_df)
'''
def empty_row_remove(data_frame):
    data_frame = data_frame.dropna(axis=1, how='all')
    if any("Unnamed" in x for x in list(data_frame)):
        data_frame.rename(columns=data_frame.iloc[0], inplace = True)
        data_frame.drop([0], inplace = True)
        return data_frame
    return data_frame
data_frame = empty_row_remove(data_frame)

def fix_column_names(data_frame):
    x = list(data_frame)
    for i in x:
        if i in part_num_names:
            data_frame.rename(columns = {i:'Part Number'}, inplace = True)
        elif "part" in i.lower().strip():
            data_frame.rename(columns={i: 'Part Number'}, inplace=True)
        elif i.strip().lower()=='part number':
            data_frame.rename(columns={i: 'Part Number'}, inplace=True)

        # distributor_number fix
        if i in distributor_num:
            data_frame.rename(columns = {i:'Distributor Number'}, inplace = True)
        elif i.strip().lower()=='distributor number':
            data_frame.rename(columns={i: 'Distributor Number'}, inplace=True)
        # reseller_names fix
        if i in reseller_names:
            data_frame.rename(columns = {i:'Reseller Name'}, inplace = True)
        elif "reseller" in i.lower().strip():
            data_frame.rename(columns={i: 'Reseller Name'}, inplace=True)
        elif i.strip().lower()=='reseller name':
            data_frame.rename(columns={i: 'Reseller Name'}, inplace=True)

        # order date
        if i in order_date_names:
            data_frame.rename(columns = {i:'Order Date'}, inplace = True)
        elif i.strip().lower()=='order date':
            data_frame.rename(columns={i: 'Order Date'}, inplace=True)
        elif "Order Date" not in x:
            if i=="Order Entry Date":
                data_frame.rename(columns={i: 'Order Date'}, inplace=True)

        # open qty name fix

        if i in open_qty_names:
            data_frame.rename(columns = {i:'Open Qty'}, inplace = True)
        quantity_strings = ["open","qty","quantity"]
        if any(x in quantity_strings for x in i.lower().strip()):
            data_frame.rename(columns={i: 'Open Qty'}, inplace=True)
        elif i.strip().lower()=='open qty':
            data_frame.rename(columns={i: 'Open Qty'}, inplace=True)


        # delivery date fix
        if i in required_date_names:
            data_frame.rename(columns = {i:'Required Delivery Date'}, inplace = True)
        elif i.strip().lower()=='required delivery date':
            data_frame.rename(columns={i: 'Required Delivery Date'}, inplace=True)

        # description fix
        if i in part_num_names:
            data_frame.rename(columns = {i:'Description'}, inplace = True)
        elif "description" in i.lower().strip():
            data_frame.rename(columns={i: 'Description'}, inplace=True)
        elif i.strip()=='Description':
            data_frame.rename(columns={i: 'Description'}, inplace=True)
    return data_frame

data_frame = fix_column_names(data_frame)

def drop_extra_columns(data_frame):

    data_frame = data_frame[cols]

    return data_frame


finished = drop_extra_columns(data_frame)
print(finished)
print(len(finished.columns))
df_list.append(finished)



'''
FILES couldnt merge:     

1. ndx_backorder_apple
Reason: Missing Dist Number Column 
'''
