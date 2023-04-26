import sys
import os
import json
import copy
import time
import os.path
import warnings
import pandas as pd
from tkinter import *
import tkinter as tk
import tkinter as ttk
from tkinter import filedialog

""" Importing this from Nandana_payout_pos_result.py """
import nandana_payout_pos_result as payout_result
# print("payout_result.... ", payout_result)
warnings.filterwarnings("ignore")
warnings.filterwarnings("ignore", category=FutureWarning)
texttoshow = None
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)


""" File path from Json """
path = open("file_paths.json")
data = json.load(path)
# print(data)

# def input1():
global contract_file
# contract_file = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls;*.csv')])
contract_file = data['contract_path']
# print("Selected Contracts file path:", contract_file)

# def input2():
global payout_files
# payout_files = filedialog.askdirectory()
payout_files = data['payout_path']
# print("Selected Payout files path:", payout_files)

# def input3():
global remittance_file
# remittance_file = filedialog.askdirectory()
remittance_file = data['remittance_path']
# print("Selected Remittance files path:", remittance_file)

# def input_4():
global filepath_4
# filepath_4 = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls;*.csv')])
filepath_4 = data['bank_path']
print("Selected Bank Statement file:", filepath_4)

# def pos():
global pos_file
# pos_file = filedialog.askdirectory()
pos_file = data['pos_path']
# print("Selected pos reports folder path is:", pos_file)

# def output():
global output_folder
# output_folder = filedialog.askdirectory()
output_folder = data['output_path']
# print("Selected output folder path is:", output_folder)

def process():
    print("Process started.")
    all_payout_lists = []
    for file in os.listdir(payout_files):
        if file.endswith('.xlsx'):
            file_path = os.path.join(payout_files, file)
            all_sheets = []
            file_merchant = file.split("_")
            for i in range(6):
                dfo = pd.read_excel(file_path, sheet_name=i, skiprows=1)
                dd = {
                    "sheet": i,
                    "dfx": dfo
                }
                # print(i, dfo[:1])
                all_sheets.append(dd)
            # print("Merchant ID from Payout files.....", file_merchant[2])
            dd = {
                "file_name": file,
                "merchant_id": file_merchant[2],
                "data_frame": all_sheets,
                "aggregator": "swiggy"
            }
            all_payout_lists.append(dd)
            # print("all_payout_lists.....", all_payout_lists)

    # contract_df = pd.read_excel("C:\\Users\\ASUS\\Downloads\\Akarsh\\Empire Contracts.xlsx")
    contract_df = pd.read_excel(contract_file)
    # print("Contract Dataframe...\n", contract_df.head())
    contract_df["M_ID"] = contract_df["Merchant ID"].astype(int)
    # print("\n Printing contract_df......", contract_df)

    # Bank Statement report Dataframe
    bank_df = pd.read_excel(filepath_4)  # Bank statement file
    print("\nBank DF Columns", bank_df.columns)
    print("\nBank DF Shape", bank_df.shape)
    bank_df = bank_df.dropna(how='all')
    bank_df = bank_df.reset_index(drop=True)
    bank_df = bank_df.drop([0, 1]).reset_index(drop=True)
    new_columns = ['Date', 'Description', 'Cheque No', 'Debit', 'Credit']
    bank_df.columns = new_columns
    bank_df['Bank Num'] = bank_df['Description'].str.split('/|-').str[-1]
    print("Bank DF Shape_aft", bank_df.shape)
    # bank_df.to_excel("Bank_DF.xlsx", index=False)


    # create a DataFrame to store the extracted data
    df = pd.DataFrame(columns=['Merchant ID', 'Aggregator', 'Date Range', 'Total Sale', 'Discounts', 'Net Bill Value', 'GST on order',
                               'Total Customer Payable', 'Commission', 'Commission Without GST', 'GST On Commission', 'Commission Percentage',
                               'Buyer cost sharing', 'Total of Order Level Adjustments', 'Other Charges And Refunds', 'HIGH_PRIORITY',
                               'Contract Details', 'Buyer Cancellation', 'TCS', 'TDS'])

    for aaa in all_payout_lists:
        merchant_id = int(aaa['merchant_id'])
        all_sheets = aaa['data_frame']
        # order_df = all_sheets[2]["dfx"]
        contract_details = contract_df[contract_df["M_ID"] == merchant_id]

        summary = all_sheets[0]["dfx"]
        invoice = all_sheets[1]["dfx"]
        order = all_sheets[2]["dfx"]
        # print("order sheet Columns.........", order.columns)

        Net_Bill = sum(order["Net Bill Value (without taxes)\nD = A + B - C"])
        GST = sum(order["GST on order (including cess)\nE"])
        payable = sum(order["Customer payable\n(Net bill value after taxes & discount)\nF = D + E"])
        Discount = sum(order["Total Merchant Discount \nC = C1 +C2"])
        TCS = sum(order["TCS\nU1"])
        TDS = sum(order["TDS\nU2"])
        col_a = sum(order["Item's total\nA"])
        col_b = sum(order["Packing & Service charges\nB"])
        total_sale = col_a + col_b
        # print("total_sale....", total_sale)
        Commission = sum(order["Total Swiggy service fee \n(including taxes)\nO = M + N"])
        # print("\n Sum of Commission... ", Commission)
        Com_without_GST = sum(order["Total Swiggy Service fee\n(without taxes)\nM = G-H+I+J+K+L"])
        # print("Sum of Commission without GST is", Com_without_GST)
        GST_Com = sum(order["Taxes on Swiggy Service fee (Including Cess)\nN"])
        # print("Sum of GST on Commission", GST_Com)
        # order["Swiggy Platform Service Fee % (%)"] = order["Swiggy Platform Service Fee % (%)"].astype(int)
        # order["Swiggy Platform Service Fee % (%)"] = order["Swiggy Platform Service Fee % (%)"].astype(float) / 100
        order["Swiggy Platform Service Fee % (%)"] = order["Swiggy Platform Service Fee % (%)"].str.replace("%", "").astype(float)/100
        # print("1st value of Commission Percentage", order.iloc[0]["Swiggy Platform Service Fee % (%)"])
        Com_per = sum(order["Swiggy Platform Service Fee % (%)"])
        # print("Commission Percentage is....", Com_per)

        Buy_Cost_sharing = sum(order["Merchant Share of Cancelled Orders\nQ = D*x%"])
        # print("Buyer_Cost_sharing.....", Buy_Cost_sharing)
        Total_order_level = sum(order["Total of Order Level Adjustments\nS = P + Q + R1 + R2 + R3"])
        # print("Total of Order Level Adjustment ", Total_order_level)
        Refund = sum(order["Refund for Customer Complaints\nR3"])
        # print("Refund.... ", Refund)
        M_Payable = sum(order["Net Payable Amount (after TCS and TDS deduction)\nV = T - U1 - U2"])
        # print("Merchant Payable....", M_Payable)

        for ds in range(len(summary)):
            if summary.iloc[ds]['Unnamed: 1'] == 'Payout Period':
                payout_range = summary.iloc[ds + 1]['Unnamed: 1']
                # print("payout_range Dates: ", payout_range)
            if "HIGH_PRIORITY" in invoice['Unnamed: 1'].values:
                Priority = invoice[invoice['Unnamed: 1'] == 'HIGH_PRIORITY']['Unnamed: 4'].values[0]
                # print("Payable:", Priority)
            if "Other Charges And Refunds" in invoice['Unnamed: 1'].values:
                Other_Charges = invoice[invoice['Unnamed: 1'] == 'Other Charges And Refunds']['Unnamed: 4'].values[0]
                # print("Payable:", Other_Charges)
        ctn_det = "NO"
        if len(contract_details.index) > 0:
            ctn_det = "YES"
            # print("getting first index values....")
            firstRow = contract_details.iloc[0]
            # print("\n contract_details........", firstRow)
            Buyer_Cancellation = firstRow["Buyer Cancellation Cost Sharing Percentage"]
            # TDS = firstRow['TDS']
            # TCS = firstRow['TCS']
            Agtr = firstRow['Aggregator']
        else:
            print("No matching contract found for Merchant ID:", merchant_id)
        print("\n Merchant ID is:", merchant_id, "contract_details:", ctn_det)

        df = df.append({'Merchant ID': merchant_id, 'Aggregator': Agtr, 'Date Range': payout_range, 'Total Sale': total_sale,
                        'Total Customer Payable': payable, 'Commission': Commission, 'Commission Without GST': Com_without_GST,
                        'GST On Commission': GST_Com, 'Commission Percentage': Com_per, 'Buyer cost sharing': Buy_Cost_sharing,
                        'Total of Order Level Adjustments': Total_order_level, 'Refund for disputed details': Refund,
                        'Other Charges And Refunds': Other_Charges, 'HIGH_PRIORITY': Priority, 'Merchant Payable': M_Payable,
                        'Net Bill Value': Net_Bill, 'Discounts': Discount, 'GST on order': GST, 'Contract Details': ctn_det,
                        'Buyer Cancellation': Buyer_Cancellation, 'TCS': TCS, 'TDS': TDS}, ignore_index=True)
        # print("df shape after.. ", df.shape)
        df = df[df['Contract Details'] == "YES"]

    # Final DF.....
    df = df.drop(columns=['Contract Details'])
    # print("final df shape... ", df.shape)

    df['Com.%'] = (df['Commission'] / df['Total Customer Payable']) * 100
    df['As per Contract'] = contract_df['Service Charges'] * 100
    df['Difference in Opinion'] = df.apply(lambda row: 'No' if row['Com.%'] >= row['As per Contract'] else 'Yes', axis=1)


    # df['Com.%'] = (df['Commission'] / df['Total Customer Payable']) * 100
    # df['As per Contract'] = contract_df['Service Charges'] * 100
    # df['Difference in Opinion'] = df.apply(lambda row: 'No' if row['Com.%'] >= row['As per Contract'] else 'Yes', axis=1)
    #
    # new_columns = ['Merchant ID', 'Aggregator', 'Date Range', 'Total Sale', 'Total Customer Payable',
    #                'Commission', 'Com.%', 'As per Contract', 'Difference in Opinion', 'Commission Without GST',
    #                'GST On Commission', 'Commission Percentage', 'Buyer cost sharing',
    #                'Total of Order Level Adjustments', 'Refund for disputed details', 'Other Charges And Refunds',
    #                'HIGH_PRIORITY', 'Merchant Payable', 'Net Bill Value', 'Discounts', 'GST on order',
    #                'Buyer Cancellation', 'TCS', 'TDS']
    # df = df.reindex(columns=new_columns)


    # # Creating DF to Excel file.....
    # df.to_excel(os.path.join(output_folder, "Swiggy Consolidated Report.xlsx"), index=False)
    # print("\nSwiggy Consolidated Report created...")

    # DF to Excel with Autofilter...
    writer = pd.ExcelWriter(os.path.join(output_folder, "Swiggy Consolidated Report.xlsx"), engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Consolidated', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Consolidated']
    # Adding autofilter
    (max_row, max_col) = df.shape
    # max_row = len(R_df.index) + 1
    # max_col = len(R_df.columns)
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column(0, max_col - 1, 15)
    writer.save()
    print("\nConsolidated Report Created.")
    time.sleep(2)

    # print("\n --------------------------------------------------------------------------------------------------------")

    # create a DataFrame to store the extracted data
    R_df = pd.DataFrame(columns=['Merchant ID', 'Outlet', 'Remarks', 'Service Range', 'Amount to be Credited', 'Date Range', 'Account Name',
                                 'Bank Number/UTR ID', 'Amount Received', 'Total Amount Received'])
    all_remittance_lists = []
    for file in os.listdir(remittance_file):
        # if file.endswith('.xlsx', '.csv'):
        if file.endswith(('.xlsx', '.csv')):
            file_path = os.path.join(remittance_file, file)
            all_sheets = []
            file_merchant = file.split("_")
            # for i in range(0):
            if file.endswith('.csv'):
                dfo = pd.read_csv(file_path)
            elif file.endswith('.xlsx'):
                dfo = pd.read_excel(file_path)
            # print("remittance_file....", remittance_file)
            # print("file_path", file_path)
            # print("file...", file)
            dd = {
                "sheet": 0,
                "dfx": dfo
            }
            # print(i, dfo[:1])
            all_sheets.append(dd)
            # print("Merchant ID from Remittance files.....", file_merchant[0])
            # print("dfo,....", dfo)
            dd = {
                "file_name": file,
                "merchant_id": file_merchant[0],
                "data_frame": all_sheets,
                "aggregator": "swiggy"
            }
            all_remittance_lists.append(dd)
            # print("all_remittance_lists.........", all_remittance_lists)
    R_df_index = 0
    for aaa in all_remittance_lists:
        mer_id = int(aaa['merchant_id'])
        contract_details = contract_df[contract_df["M_ID"] == mer_id]
        Outlet_Name = " "
        if len(contract_details.index) > 0:
            firstRow = contract_details.iloc[0]
            Outlet_Name = firstRow["Outlet Name"]

        summary = aaa['data_frame'][0]["dfx"]
        summary['Outlet Name'] = Outlet_Name
        Amt_sum = summary['Amount'].sum()
        # print("Summary.........", summary)

        for index, row in summary.iterrows():
            R_df.loc[R_df_index, 'Merchant ID'] = row['Merchant Id']
            R_df.loc[R_df_index, 'Outlet'] = row["Outlet Name"]
            R_df.loc[R_df_index, 'Service Range'] = row["Service Period"]
            R_df.loc[R_df_index, 'Account Name'] = row["Account"]
            R_df.loc[R_df_index, 'Date Range'] = row["Date"]
            R_df.loc[R_df_index, 'Bank Number/UTR ID'] = row["Bank Reference Number"]
            # R_df.loc[R_df_index, 'Amount to be Credited'] = Amt_sum
            R_df.loc[R_df_index, 'Amount Received'] = row["Amount"]
            # R_df.loc[R_df_index, 'Total Amount Received'] = Amt_sum
            R_df_index += 1
            # Only set total amount received for the first row with this merchant ID
            if index == summary.index[0]:
                R_df.loc[R_df_index, 'Total Amount Received'] = Amt_sum
                R_df.loc[R_df_index, 'Amount to be Credited'] = Amt_sum

    # R_df['Remarks'] = R_df['Date Range'] + ', ' + R_df['Bank Number/UTR ID'] + ', ' + R_df['Amount Received']
    # R_df['Remarks'] = R_df['Date Range'] + ', ' + R_df['Bank Number/UTR ID'] + ', ' + R_df['Amount Received'].astype(str)
    R_df['Remarks'] = 'Date Range: ' + R_df['Date Range'] + ', Bank Number/UTR ID: ' + R_df[
        'Bank Number/UTR ID'] + ', Amount Received: ' + R_df['Amount Received'].astype(str)


    # Bank number & credit Amt
    bank_df['Bank Num'] = bank_df['Bank Num'].str.lower()
    R_df['Bank Number/UTR ID'] = R_df['Bank Number/UTR ID'].str.lower()

    # merge the two dataframes on the matching columns
    merged_df = pd.merge(R_df, bank_df, left_on='Bank Number/UTR ID', right_on='Bank Num', how='left')
    # merged_df.to_excel("Merged_DF.xlsx", index=False)
    # create new column with "Credit" information from bank_df in R_df
    R_df['As per Bank'] = merged_df['Credit']


    # rows = [["Remarks"], ['']]
    # cols = [['Remarks', 'Date Range', 'Bank Number/UTR ID', 'Amount Received'], ['Remarks', '2'], ['Remarks', '3'], ['Remarks', '4']]
    # workbook = xlsxwriter.Workbook('Swiggy.xlsx')
    # worksheet = workbook.add_worksheet()
    # for i, row in enumerate(rows):
    #     for j, col in enumerate(cols):
    #         for el in row:
    #             if el in col:
    #                 worksheet.write(i, j, el)
    # workbook.close()


    # # DF to Excel...
    # R_df.to_excel(os.path.join(output_folder, "Swiggy Remittance Report.xlsx"), index=False)
    # print("Swiggy Remittance Report created...")
    # label_file_explorer.configure(text="Process Completed & files saved in Output folder")

    # DF to Excel with Autofilter...
    writer = pd.ExcelWriter(os.path.join(output_folder, "Swiggy Remittance Report.xlsx"), engine='xlsxwriter')
    R_df.to_excel(writer, sheet_name='Remittance', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Remittance']
    # Add autofilter
    (max_row, max_col) = R_df.shape
    # max_row = len(R_df.index) + 1
    # max_col = len(R_df.columns)
    worksheet.set_column(0, max_col - 1, 20)
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    writer.save()
    print("Remittance Report Created.")
    time.sleep(2)

    # print("--------------------------------------------- Nandana POS -----------------------------------------------")

    def update_window(texttoshow, window, msg):
        try:
            window.update_idletasks()
            window.update()
            texttoshow.set(msg)
        except:
            pp = 0
    project_path = os.getcwd()
    if output_folder:
        dir_name = []
        dir_name.append(output_folder)
        # print("Output path imported")

    if payout_files:
        dir_name_payout = []
        dir_name_payout.append(payout_files)
        # print("Payout path imported ")

    if pos_file:
        dir_name_pos = []
        dir_name_pos.append(pos_file)
        # print("Pos folder path... ")

        pos_dir_path = dir_name_pos[0]
        payout_dir_path = dir_name_payout[0]

        op_dir_path = dir_name[0]
        try:
            payout_result.nandana_get_output(contract_file, pos_dir_path, payout_dir_path, op_dir_path)
            # label_1.configure(text="Process Completed")
            print("POS Order file created")
        except Exception as e:
            exception_type, exception_object, exception_traceback = sys.exc_info()
            contract = exception_traceback.tb_frame.f_code.co_filename
            only_filename = contract.split("\\")[-1].replace(".py", "")
            line_number = exception_traceback.tb_lineno
            # print("+++ Type >>", exception_type)
            # print("+++ File >>", only_filename)
            # print("+++ Line >>", line_number)
            # print("+++ Error >>", e)
            # print(e)
            # print("ONEPAGE_WORD", traceback.format_exc())
            lines = [f"+++ Type >> {exception_type}", f"+++ File >> {only_filename}", f"+++ Line >> {line_number}",
                     f"+++ Error >> {e}", ]
            # print(lines)
            with open(f'{project_path}\\errorLog.txt', 'a') as f:
                f.write('#' * 25)
                f.write('main file')
                f.write('\n')

                for line in lines:
                    f.write(str(line))
                    f.write('\n')
                f.write('\n')
                f.write('#' * 25)

            # update_window(file_status, root, 'Error - please check the logs')
            label_1.configure(text="Error - please check the logs")
    dir_name = []
    dir_name_pos = []
    dir_name_payout = []

    # print(--------------------------------------- Discount Part----------------------------------------------)
    flext = contract_file.split('.')[-1]
    if flext == 'csv':
        contracts_dataframe = pd.read_csv(f'{contract_file}', engine='openpyxl')
    else:
        contracts_dataframe = pd.read_excel(f'{contract_file}', engine='openpyxl')

    def get_dif_rec(df_columns=None, dis_app=None, mer_share=None, dif_rec=None):
        rows_data = copy.deepcopy(df_columns)
        dis_app_value = float(rows_data[dis_app])
        mer_share_val = float(rows_data[mer_share])
        rows_data[dif_rec] = 0
        if (mer_share_val > 0 and dis_app_value > mer_share_val):
            rows_data[dif_rec] = dis_app_value - mer_share_val
        return rows_data

    def get_dif_opt(val):
        rval = "No"
        if (val > 0):
            rval = "Yes"
        return rval

    def get_dif_dis(val):
        rval = ""
        if (val > 0):
            rval = "Upto " + str(val) + '%'
        return rval

    for df in range(len(contracts_dataframe)):
        merch_id = str(contracts_dataframe.iloc[df]['Merchant ID'])
        for file in os.listdir(payout_dir_path):
            filenames = file.split('_')
            if merch_id in filenames:
                payout_ext = file.split('.')[-1]
                new_filepath = payout_dir_path + "\\" + file
                if payout_ext == 'csv':
                    pay_all = pd.read_csv(f'{payout_dir_path}\\{file}', sheet_name='All Orders')
                else:
                    pay_all = pd.read_excel(f'{payout_dir_path}\\{file}', sheet_name='All Orders')
                # print("\nPayout all_order sheet DF shape...", pay_all.shape)
                # print("\n Payout all_order sheet columns...", pay_all.columns)
                if payout_ext == 'csv':
                    pay_PL = pd.read_csv(f'{payout_dir_path}\\{file}', sheet_name='Discounts P&L', skiprows=7)
                else:
                    pay_PL = pd.read_excel(f'{payout_dir_path}\\{file}', sheet_name='Discounts P&L', skiprows=7)
                # print("File..", file)
                # print("File path", payout_dir_path)
                # print("Payout Discount P&L sheet DataFrame", pay_PL.shape)
                # print("\n first row...", pay_PL)
                col_index = 0
                col_list = []
                for i in pay_PL.iloc[0]:
                    # print(i)
                    if str(i) == "nan":
                        col_name = "temp" + str(col_index)
                    else:
                        col_name = str(i).replace("\n", " ").strip()
                    col_list.append(col_name)
                    col_index += 1
                pay_PL.columns = col_list
                pay_PL = pay_PL.drop(pay_PL.columns[0], axis=1)

                # removing empty values from Merchant Share col
                pay_PL = pay_PL.dropna(subset=["Merchant Share (%)"])
                pay_PL["Merchant Share (%)"] = pay_PL["Merchant Share (%)"].astype(str).str.strip()
                pay_PL = pay_PL[pay_PL["Merchant Share (%)"].apply(lambda x: x.isdigit())]
                pay_PL = pay_PL.dropna(how='all')  # removing empty rows

                # Updating the mssing rows with previous value
                pay_PL["Campaign ID"] = pay_PL["Campaign ID"].fillna(method="ffill")
                pay_PL["Validity"] = pay_PL["Validity"].fillna(method="ffill")
                # print("\npay_PL shape after updating", pay_PL.shape)
                # pay_PL.to_excel("Pay_PL(filtered).xlsx", index=False)

                pattern = r"Use code (\w+) & get (\d+%)\s+off on orders above ₹(\d+)\. Maximum discount: ₹(\d+)"
                extracted = pay_PL["Description"].str.extract(pattern)
                extracted.columns = ["Code", "Discount", "Min Order Amount", "Max Discount"]
                pay_PL = pay_PL.join(extracted)
                # print("PAy PL.......", pay_PL)
                pay_PL["Unique Value"] = pay_PL["Campaign ID"].astype(str) + "_" + pay_PL["Validity"].astype(
                    str) + "_" + \
                                         pay_PL["Coupon"].astype(str) + "_" + pay_PL["Merchant Share (%)"].astype(
                    str) + "_" + \
                                         pay_PL["Description"].astype(str)
                pay_PL["Unique Value"] = pay_PL["Unique Value"].apply(lambda x: '_'.join(list(set(x.split('_')))))
                pay_PL["Swiggy Share %"] = pay_PL["Merchant Share (%)"].apply(lambda x: 100 - float(x))
                pay_PL = pay_PL[['Campaign ID', 'Coupon', "Code", "Discount", "Min Order Amount", "Max Discount",
                                 "Merchant Share (%)", 'Swiggy Share %', 'Description']]
                pay_PL = pd.DataFrame(pay_PL)
                # pay_PL.to_excel("Pay_PL_final.xlsx", index=False)

                headers = pay_all.iloc[0]
                # print("Headers...", headers)
                new_payout_dataframe = pd.DataFrame(pay_all.values[1:], columns=headers)
                column_names = ["Restaurant ID", "Date of order", "Campaign ID", "Order ID", "Order Value",
                                "Order Value Discount", "Agreed Discount", "Net Value in Payout",
                                "Net Value in Discount", "Difference Recorded", "As per agreed", "As actual Recorded",
                                "Difference in Opinion", "Ref Pay out path", "Coupon code applied by customer",
                                "GST on order (including cess)\nE",
                                'Total Merchant Discount \nC = C1 +C2']

                result_dataframe = pd.DataFrame(columns=column_names)
                result_data = []
                result_data.append(column_names)

                new_payout_dataframe.loc[new_payout_dataframe[
                                             "Total Merchant Discount \nC = C1 +C2"] == "", "Total Merchant Discount \nC = C1 +C2"] = 0
                new_payout_dataframe.loc[new_payout_dataframe[
                                             "Total Merchant Discount \nC = C1 +C2"] == " ", "Total Merchant Discount \nC = C1 +C2"] = 0
                new_payout_dataframe["Total Merchant Discount \nC = C1 +C2"] = new_payout_dataframe[
                    "Total Merchant Discount \nC = C1 +C2"].fillna(0).astype(float)

                new_payout_dataframe.loc[new_payout_dataframe[
                                             "GST on order (including cess)\nE"] == "", "GST on order (including cess)\nE"] = 0
                new_payout_dataframe.loc[new_payout_dataframe[
                                             "GST on order (including cess)\nE"] == " ", "GST on order (including cess)\nE"] = 0
                new_payout_dataframe["GST on order (including cess)\nE"] = new_payout_dataframe[
                    "GST on order (including cess)\nE"].fillna(0).astype(float)

                for disc_id in pay_PL['Campaign ID']:
                    rslt_df = new_payout_dataframe[new_payout_dataframe['Discount Campaign ID'] == str(disc_id)]
                    head_log = {"Restaurant ID": [' '], "Date of order": [' '], "Campaign ID": [' '],
                                "Order ID": [' '], "Order Value": [' '], "Order Value Discount": [' '],
                                "Net Value in Payout": [' '],
                                "Net Value in Discount": [' '], "Difference Recorded": [' '], "As per agreed": [' '],
                                "As actual Recorded": [' '], "Difference in Opinion": [' '], "Ref Pay out path": [' '],
                                "Coupon code applied by customer": [' '], "As per Payout": [' '], "Swiggy Share": [' '],
                                "GST Amt": [' ']}
                    # result_data.append(["", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "])
                    head_df = pd.DataFrame(head_log)
                    headframes = [result_dataframe, head_df]

                    result_dataframe = pd.concat(headframes)

                    for payout_row in range(len(rslt_df)):

                        row = rslt_df.iloc[payout_row]
                        order_date = row["Order Date"]
                        order_no = row["Order No"]
                        # items_total = row["Item's total\nA"] + row["Packing & Service charges\nB"]
                        items_total = row["Item's total\nA"]
                        discount = row["Merchant Discount\nC1"] + row["Exclusive Offer\nC2"]
                        net_bill_value = row["Net Bill Value (without taxes)\nD = A + B - C"]

                        Customer_coupon_code = row["Coupon code applied by customer"]
                        discount_type_id = row["Discount Campaign ID"]
                        net_bill_value_discount = abs(items_total - discount)
                        net_value_diffrence = abs(net_bill_value_discount - net_bill_value)

                        As_per_Payout = row["Total Merchant Discount \nC = C1 +C2"]
                        GST_AMT = row["GST on order (including cess)\nE"]

                        if int(net_value_diffrence) == 0:
                            difference_opinion = 'No'
                        else:
                            difference_opinion = 'Yes'

                        df_log = {"Restaurant ID": [str(merch_id)], "Date of order": [order_date],
                                  "Campaign ID": [disc_id],
                                  "Order ID": [order_no], "Order Value": [items_total],
                                  "Order Value Discount": [discount],
                                  "Net Value in Payout": [net_bill_value],
                                  "Net Value in Discount": [net_bill_value_discount],
                                  "Difference Recorded": [net_value_diffrence],
                                  "As per agreed": [disc_id], "As actual Recorded": [disc_id],
                                  "Difference in Opinion": [difference_opinion], "Ref Pay out path": [file],
                                  "Coupon code applied by customer": [Customer_coupon_code],
                                  "As per Payout": [As_per_Payout], "GST Amt": [GST_AMT]}

                        result_data.append([str(merch_id), order_date, disc_id, order_no, items_total, discount, net_bill_value,
                             net_bill_value_discount, net_value_diffrence, disc_id, disc_id, difference_opinion, file,
                             Customer_coupon_code, As_per_Payout, GST_AMT])
                        row_df = pd.DataFrame(df_log)
                        frames = [result_dataframe, row_df]
                        result_dataframe = pd.concat(frames)
                        result_dataframe['Campaign ID'] = result_dataframe['Campaign ID'].astype(str)
                        result_dataframe['Coupon'] = result_dataframe['Coupon code applied by customer'].astype(str)

                    """ Matching Campaign ID & Coupon code in both DF """

                    final_dataframe = pd.merge(result_dataframe, pay_PL, how="left", on=["Campaign ID", "Coupon"])
                    final_dataframe = final_dataframe[final_dataframe['Restaurant ID'].notna()]
                    final_dataframe = final_dataframe[final_dataframe['Restaurant ID'] != ""]
                    final_dataframe = final_dataframe[final_dataframe['Restaurant ID'] != " "]
                    final_dataframe['Total Discount'] = 0

                    final_dataframe.loc[final_dataframe["Order Value"] == "", "Order Value"] = 0
                    final_dataframe.loc[final_dataframe["Order Value"] == " ", "Order Value"] = 0
                    final_dataframe["Order Value"] = final_dataframe["Order Value"].fillna(0).astype(float)

                    final_dataframe.loc[final_dataframe["Order Value Discount"] == "", "Order Value Discount"] = 0
                    final_dataframe.loc[final_dataframe["Order Value Discount"] == " ", "Order Value Discount"] = 0
                    final_dataframe["Order Value Discount"] = final_dataframe["Order Value Discount"].fillna(0).astype(float)

                    final_dataframe['Total Discount'] = round(((final_dataframe['Order Value'].astype(float) *
                                                                final_dataframe['Discount'].str.replace('%', '').astype(float)) / 100), 2)

                    # final_dataframe.loc[final_dataframe["Merchant Cost Sharing"] == "", "Merchant Cost Sharing"] = 0
                    final_dataframe.loc[final_dataframe["Merchant Share (%)"] == "", "Merchant Share (%)"] = 0
                    final_dataframe["Merchant Share (%)"] = final_dataframe["Merchant Share (%)"].fillna(0).astype(int)
                    final_dataframe["Difference Recorded"] = ""
                    final_dataframe["Difference in Opinion"] = ""
                    dis_app = "Total Discount"
                    mer_share = "Merchant Share (%)"
                    dif_rec = "Difference Recorded"

                    final_dataframe = final_dataframe.apply(lambda x: get_dif_rec(x, dis_app, mer_share, dif_rec),
                                                            axis=1)
                    final_dataframe["Difference in Opinion"] = final_dataframe["Difference Recorded"].apply(get_dif_opt)
                    final_dataframe["Agreed Discount"] = final_dataframe["Merchant Share (%)"].apply(get_dif_dis)

                    final_dataframe['Total Discount'] = final_dataframe['Total Discount'].astype(float)
                    final_dataframe['Max Discount'] = final_dataframe['Max Discount'].astype(float)
                    final_dataframe['Discount Applied Amount'] = final_dataframe.apply(
                        lambda row: min(row['Total Discount'], row['Max Discount']), axis=1)

                    final_dataframe['Merchant Share (%)'] = final_dataframe['Merchant Share (%)'].astype(str)
                    final_dataframe['Merchant Share Value'] = round(
                        ((final_dataframe['Discount Applied Amount'].astype(float) *
                          final_dataframe['Merchant Share (%)'].str.replace('%', '').astype(float)) / 100), 2)

                    final_dataframe['Difference in Opinion'] = (
                                final_dataframe['As per Payout'].astype(float) - final_dataframe['Merchant Share Value']
                                .astype(float)).ne(0).map({True: 'Yes', False: 'No'})

                    final_dataframe['Swiggy Share'] = round(((final_dataframe['Discount Applied Amount'].astype(float) *
                                                              final_dataframe['Swiggy Share %'].astype(str).str.replace(
                                                                  '%', '').astype(float)) / 100), 2)

                    # Dropping the repeated values from Order ID colun..
                    final_dataframe.drop_duplicates(subset=['Order ID'], inplace=True)

                    final_dataframe.drop(
                        ['Agreed Discount', 'Net Value in Payout', 'Net Value in Discount', 'Difference Recorded',
                         'As per agreed', 'As actual Recorded', 'Coupon code applied by customer',
                         'Description', 'Ref Pay out path', 'Order Value Discount'], axis=1, inplace=True)

                    final_dataframe = final_dataframe.loc[:,
                                      ['Restaurant ID', 'Date of order', 'Campaign ID', 'Order ID', 'Coupon',
                                       'Order Value',
                                       'Discount', 'Min Order Amount', 'Max Discount', 'Total Discount',
                                       'Discount Applied Amount',
                                       'Merchant Share (%)', 'Swiggy Share %', 'Merchant Share Value',
                                       'Difference in Opinion', 'As per Payout',
                                       'Swiggy Share']]
                    final_dataframe.to_excel(f"{op_dir_path}\\{merch_id} Discount_Report_test.xlsx", index=False)
    print("Discount reports created.")

    # ------------------------------------------- Cancellation -------------------------------------
    flext = contract_file.split('.')[-1]
    if flext == 'csv':
        contracts_dataframe = pd.read_csv(f'{contract_file}', engine='openpyxl')
    else:
        contracts_dataframe = pd.read_excel(f'{contract_file}', engine='openpyxl')

    for df in range(len(contracts_dataframe)):

        merch_id = str(contracts_dataframe.iloc[df]['Merchant ID'])

        for file in os.listdir(payout_files):

            filenames = file.split('_')
            if merch_id in filenames:
                payout_ext = file.split('.')[-1]

                if payout_ext == 'csv':
                    payout_dataframe = pd.read_csv(f'{payout_files}\\{file}', sheet_name='All Orders')
                else:
                    payout_dataframe = pd.read_excel(f'{payout_files}\\{file}', sheet_name='All Orders')

                headers = payout_dataframe.iloc[0]
                new_payout_dataframe = pd.DataFrame(payout_dataframe.values[1:], columns=headers)

                rslt_df = new_payout_dataframe[new_payout_dataframe['Order Status'] == str('cancelled')]

                # "Difference in Opinion" ,	"Differnce Recorded"
                column_names = ["Restaurant ID", "Date of order", "Date & Time of Cancellation", "Order ID",
                                "Item Total", "Packing Chargers", "Discounts",
                                "Net Value", "GST", "Order Value", "Cancelation Value", "GST on Value",
                                "Cancelation Value Captured", "As per policy", "As per policy Captured", \
                                "As actual Recorded", "Ref Pay out path"]

                result_dataframe = pd.DataFrame(columns=column_names)

                for payout_row in range(len(rslt_df)):
                    row = rslt_df.iloc[payout_row]
                    items_total = row["Item's total\nA"]
                    if int(float(items_total)) != 0:
                        order_date = row["Order Date"]
                        cancellation_time = str(row["Order Date"]) + ' ' + str(row["Cancellation time"])
                        order_no = row["Order No"]

                        packing_charge = row["Packing & Service charges\nB"]
                        discount = row["Total Merchant Discount \nC = C1 +C2"]

                        net_bill_value = row["Net Bill Value (without taxes)\nD = A + B - C"]
                        gst_on_order = row["GST on order (including cess)\nE"]
                        customer_payble = row["Customer payable\n(Net bill value after taxes & discount)\nF = D + E"]
                        merch_share = row["Merchant Share of Cancelled Orders\nQ = D*x%"]
                        gst_deduct = row["GST Deduction U/S 9(5)\nR2"]
                        cancellation_value = row["Total of Order Level Adjustments\nS = P + Q + R1 + R2 + R3"]
                        cancellation_policy = row["Cancellation Policy Applied"]
                        pickup_status = row["Pick Up Status"]

                        cancelled_by = row["Cancelled By?"]

                        df_log = {"Restaurant ID": [str(merch_id)], "Date of order": [order_date],
                                  "Date & Time of Cancellation": [cancellation_time],
                                  "Order ID": [order_no], "Item Total": [items_total],
                                  "Packing Chargers": [packing_charge], "Discounts": [discount],
                                  "Net Value": [net_bill_value], "GST": [gst_on_order],
                                  "Order Value": [customer_payble], "Cancelation Value": [merch_share],
                                  "GST on Value": [gst_deduct], "Cancelation Value Captured": [cancellation_value],
                                  "As per policy": [cancellation_policy],
                                  "As per policy Captured": [pickup_status],
                                  "As actual Recorded": [cancelled_by], "Ref Pay out path": [file]}

                        row_df = pd.DataFrame(df_log)
                        frames = [result_dataframe, row_df]

                        result_dataframe = pd.concat(frames)

                # result_dataframe.to_excel(f'{op_dir_path}\\Cancellation Report.xlsx', sheet_name='Cancellation Report', index=False)

                # cancellation_dataframe = pd.read_excel(f'{op_dir_path}\\Cancellation Report.xlsx')
                cancellation_dataframe = result_dataframe
                cat_4 = [str(x) for x in cancellation_dataframe["Order ID"].tolist()]
                # cat_4 = ['Metric_' + str(x) for x in range(len(cancellation_dataframe))]
                # index_4 = ['Data 1', 'Data 2', 'Data 3', 'Data 4']
                index_4 = ["Order Value", "Cancelation Value", "GST on Value", "Cancelation Value Captured"]
                # print('cat4' , cat_4 , '----' , 'index1' ,index_4)

                data_3 = {}
                for cat in range(len(cat_4)):
                    data_3[cat_4[cat]] = [(cancellation_dataframe["Order Value"]).tolist()[cat],
                                          cancellation_dataframe["Cancelation Value"].tolist()[cat],
                                          cancellation_dataframe["GST on Value"].tolist()[cat],
                                          cancellation_dataframe["Cancelation Value Captured"].tolist()[cat]]
                # print('data3' , data_3)
                # Create a Pandas dataframe from the data.

                df = pd.DataFrame(data_3, index=index_4)

                # Create a Pandas Excel
                # writer using XlsxWriter as the engine.
                excel_file = f'{output_folder}\\{merch_id} Cancellation_report_final.xlsx'
                sheet_name = ['Cancellation Report', 'Sheet2']

                writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
                result_dataframe.to_excel(writer, sheet_name=sheet_name[0])
                df.to_excel(writer, sheet_name=sheet_name[1])

                # Access the XlsxWriter workbook and worksheet objects from the dataframe.
                workbook = writer.book
                worksheet = writer.sheets[sheet_name[0]]

                # Create a chart object.
                chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

                # Configure the series of the chart from the dataframe data.
                # for col_num in range(1, len(cat_4) + 1):
                #     chart.add_series({
                #         'name':       ['Sheet1', 0, col_num],
                #         'categories': ['Sheet1', 1, 0, 4, 0],
                #         'values':     ['Sheet1', 1, col_num, 4, col_num],
                #         'gap':        2,
                #     })
                #
                for col_num in range(1, len(cat_4) + 1):
                    chart.add_series({
                        'name': ['Sheet2', 0, col_num],
                        'categories': ['Sheet2', 1, 0, 4, 0],
                        'values': ['Sheet2', 1, col_num, 4, col_num],
                        'gap': 2,
                    })
                # Configure the chart axes.
                chart.set_y_axis({'major_gridlines': {'visible': False}})
                # Insert the chart into the worksheet.
                worksheet.insert_chart('A2', chart)
                # Close the Pandas Excel writer and output the Excel file.
                writer.save()
                label_1.configure(text="Process completed and all Files saved in Output folder")
                time.sleep(3)
    print("Cancellation report created.")
    print("\nAll Process completed")




root = tk.Tk()
root.title(" Mindful Automation Private Limited")
root.geometry("400x130")
root['bg'] = 'white'
current_path = os.path.dirname(os.path.realpath(__file__))
root.wm_iconbitmap(f"{current_path}/icons/mindful_logo.ico")

# Create a canvas & adding scrollbar
canvas = tk.Canvas(root, bg="white")
canvas.pack(side="left", fill="both", expand=True)
scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
frame = tk.Frame(canvas, bg="white")
canvas.create_window((0, 0), window=frame, anchor="nw")

label_1 = tk.Label(frame, text="F&B Report Automation.", width=60, height=3, fg="#03001C")
label_1.grid(column=1, row=1)
# input_file_1 = tk.Button(frame, text="Select the Contract Excel file", command=input1, height=1, width=32)
# input_file_1.grid(column=1, row=2, pady=10)
# input_file_2 = tk.Button(frame, text="Select the Payout files/folder", command=input2, height=1, width=32)
# input_file_2.grid(column=1, row=3, pady=10)
# input_file_3 = tk.Button(frame, text="Select the Remittance files/folder", command=input3, height=1, width=32)
# input_file_3.grid(column=1, row=4, pady=10)
# input_file_4 = tk.Button(frame, text="Select Discount Campin ID Excel file", command=input4, height=1, width=32)
# input_file_4.grid(column=1, row=5, pady=10)
# input_file_5 = tk.Button(frame, text="Select POS folder/file", command=pos, height=1, width=32)
# input_file_5.grid(column=1, row=6, pady=10)
# output_folder = tk.Button(frame, text="Select the Output folder", command=output, height=1, width=32)
# output_folder.grid(column=1, row=7, pady=10)
start_button = tk.Button(frame, text='Start Process', command=process, height=1, width=40)
start_button.grid(column=1, row=2, pady=10)
global file_status
file_status = StringVar()
file_status.set('')
root.mainloop()

