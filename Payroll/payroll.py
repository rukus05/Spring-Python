



from msilib.schema import File
import time
import pandas as pd
import re
import openpyxl
import datetime
import tkinter as tk
from tkinter import TOP, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.filedialog import asksaveasfile
import PySimpleGUI as sg


## Allocations.py runs first.
## CreatePivot.py runs second.
## To generate the PAYROLL, run this File first.  Execute the Allocations.bat file.
## The output created from this file is ingested into the CreatePivot.py file.
def main(): 
    
    start = time.time()
    # Prompt user for the Raw data

    f = FilePrompt()
    df_spring = pd.read_excel(f)

    
    df_spring = df_spring.reset_index()
    
    # Get the unique Invoice Numbers, Locations, Sub Departments, and Department Long Descriptions.  This is needed to loop through and group them properly.
    uniqueInvoices = df_spring['Invoice Number'].unique()
    uniqueLocations = df_spring['LOCATION'].unique()
    uniqueSub_Dept = df_spring['SUB_DEPARTMENT'].unique()
    unique_DLD = df_spring['Department Long Descr'].unique()
    #print(uniqueInvoices)
    
    #  It's important to fill in blank cells for the below columns with Zeros.  A blank cell breaks the calculations
    #  The .fillna() method fills blank cells in these columns with 0's.
    df_spring['Gross Wages'] = df_spring['Gross Wages'].fillna(0)
    df_spring['OT'] = df_spring['OT'].fillna(0)
    df_spring['Bonus'] = df_spring['Bonus'].fillna(0)
    df_spring['Taxes - ER - Totals'] = df_spring['Taxes - ER - Totals'].fillna(0)
    df_spring['Workers Comp Fee - Totals'] = df_spring['Workers Comp Fee - Totals'].fillna(0)
    df_spring['401k/Roth-ER'] = df_spring['401k/Roth-ER'].fillna(0)
    df_spring['BENEFITS wo 401K'] = df_spring['BENEFITS wo 401K'].fillna(0)
    df_spring['TOTAL FEES'] = df_spring['TOTAL FEES'].fillna(0)
    df_spring['PTO2'] = df_spring['PTO2'].fillna(0)
    df_spring['Electronics Nontaxable'] = df_spring['Electronics Nontaxable'].fillna(0)
    df_spring['Reimbursement-Non Taxable'] = df_spring['Reimbursement-Non Taxable'].fillna(0)
    df_spring['Total Client Charges'] = df_spring['Total Client Charges'].fillna(0)
    
    # Create new Dataframe for the Exceptions Output.
    df_exceptions = pd.DataFrame(columns=['Employee Name', 'Invoice Number', 'Pay End Date', 'Invoice Date', 'LOCATION', 'SUB_DEPARTMENT', 'Department Long Descr', 'DEPT CODE', 'Gross Wages', 'OT', 'Bonus', 'Taxes - ER - Totals', 'Workers Comp Fee - Totals', '401k/Roth-ER', 'BENEFITS wo 401K', 'TOTAL FEES', 'PTO2', 'Electronics Nontaxable', 'Reimbursement-Non Taxable', 'Total Client Charges'])
    
    # Create new Dataframe for the Output.
    df_Output = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', 'DebitAmt', 'CreditAmt', 'Loc', 'Dept', 'Provider', 'Service Line', 'Comments'])
    #
    exc_Dict = {}   # Exclusion Dict
    cc_Dict = {}    # Call Center Dict
    mr_Dict = {}    # Medical Records Dict
    rhq_Dict = {}   # Reception HQ Dict
    fc_Dict = {}    # Financial Counselor Dict
    co_Dict = {}    # Clinical Operations Dict
    SFOAKSV = ['SF', 'OAK', 'SV']
    SFHQSV = ['SF', 'HQ', 'SV']
    SFOAKSVNYC = ['SF', 'OAK', 'SV', 'NYC']
    SFOAKSVNYCNEST = ['SF', 'OAK', 'SV', 'NYC', 'Nest']
    HQNEST = ['HQ', 'Nest']
    SFSV = ['SF', 'SV']
    NYCMSO = ['NYC', 'MSO']
    dldloc = df_spring.columns.get_loc('Department Long Descr')
    locloc = df_spring.columns.get_loc('LOCATION')
    # First group of 4 For loops is to Handle (Clean) Exceptions
    
    # Reclassify some exceptions before massaging the dataframe
    # Add all People to the appropriate Dict Data Structure:
    # Call Center, Medical Records, Receptionist HQ, Financial Counselor, Clincal Operations
    for index, row in df_spring.iterrows():
        #if (row['Department Long Descr'] == 'Call Center') and (row['Employee Name'] != 'Lee,Stephannie Victoria'): 
        if re.match('Call Center*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Call Center '):
            cc_Dict[row['Employee Name']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
        if re.match('Medical Records*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Medical Records'): 
            mr_Dict[row['Employee Name']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
        if re.match('Receptionist HQ*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Receptionist HQ'): 
            rhq_Dict[row['Employee Name']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
        if re.match('Financial Counselor*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Financial Counselor'): 
            fc_Dict[row['Employee Name']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
        if re.match('Clinical Operations*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Clinical Operations'): 
            co_Dict[row['Employee Name']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
    
    
    #print(fc_Dict)

    for i in uniqueInvoices:
        for k in uniqueSub_Dept:
            for j in unique_DLD:
                
                # Handle Exceptions
                """
                #Audrey Krall
                exc_Dict[2495811] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # Dean, Ursula
                exc_Dict[1906920] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # Phuong Dam
                exc_Dict[2167864] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                 # My Dung Lee
                exc_Dict[1788698] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # Trieu,Minh Hue
                exc_Dict[1796892] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # Lagano,Lauren
                exc_Dict[10328183] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # Bell,Allie
                exc_Dict[1792388] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # Tran,Nam D
                exc_Dict[1804091] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # Klatsky,Peter
                exc_Dict[1770296] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                #Sergio Vaccari
                exc_Dict[1785954] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                """
                exc_Dict["Krall,Audrey"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Dean,Ursula J"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # exc_Dict["Dam,Phuong My"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # exc_Dict["Lee,My Dung"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # exc_Dict["Trieu,Minh Hue"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # exc_Dict["Lagano,Lauren"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                # exc_Dict["Bell,Allie"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Tran,Nam D"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Klatsky,Peter"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Vaccari,Sergio"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Spivey,Allison Michele"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]

                ##
                ##  The below for loops may be redundant--> logic executed in rows 92 & 95

                for key in cc_Dict:
                    cc_Dict[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                for key in mr_Dict:
                    mr_Dict[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                for key in rhq_Dict:
                    rhq_Dict[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                for key in fc_Dict:
                    fc_Dict[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                for key in co_Dict:
                    co_Dict[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]

                for index, row in df_spring.iterrows():
                    if (row['Invoice Number'] == i) and (row['Department Long Descr'] == j) and (row['SUB_DEPARTMENT'] == k):
                        
                        for key in exc_Dict:
                            if (row['Employee Name'] == key):
                                exc_Dict[key][0] = exc_Dict[key][0] + row['Gross Wages']
                                exc_Dict[key][1] = exc_Dict[key][1] + row['OT']
                                exc_Dict[key][2] = exc_Dict[key][2] + row['Bonus']
                                exc_Dict[key][3] = exc_Dict[key][3] + row['Taxes - ER - Totals']
                                exc_Dict[key][4] = exc_Dict[key][4] + row['Workers Comp Fee - Totals']
                                exc_Dict[key][5] = exc_Dict[key][5] + row['401k/Roth-ER']
                                exc_Dict[key][6] = exc_Dict[key][6] + row['BENEFITS wo 401K']
                                exc_Dict[key][7] = exc_Dict[key][7] + row['TOTAL FEES']
                                exc_Dict[key][8] = exc_Dict[key][8] + row['PTO2']
                                exc_Dict[key][9] = exc_Dict[key][9] + row['Electronics Nontaxable']
                                exc_Dict[key][10] = exc_Dict[key][10] + row['Reimbursement-Non Taxable']
                                exc_Dict[key][11] = exc_Dict[key][11] + row['Total Client Charges']
                                exc_Dict[key][12] = row['DEPT CODE']
                                exc_Dict[key][13] = row['Employee Name']
                                exc_Dict[key][14] = row['Pay End Date']
                                exc_Dict[key][15] = row['Invoice Date']
                                exc_Dict[key][16] = True
                                df_spring = df_spring.drop(index)

                        for key2 in cc_Dict:
                            if (row['Employee Name'] == key2):
                                #print(key2, " ", i, " ", j, " ", k)
                                cc_Dict[key2][0] = cc_Dict[key2][0] + row['Gross Wages']
                                cc_Dict[key2][1] = cc_Dict[key2][1] + row['OT']
                                cc_Dict[key2][2] = cc_Dict[key2][2] + row['Bonus']
                                cc_Dict[key2][3] = cc_Dict[key2][3] + row['Taxes - ER - Totals']
                                cc_Dict[key2][4] = cc_Dict[key2][4] + row['Workers Comp Fee - Totals']
                                cc_Dict[key2][5] = cc_Dict[key2][5] + row['401k/Roth-ER']
                                cc_Dict[key2][6] = cc_Dict[key2][6] + row['BENEFITS wo 401K']
                                cc_Dict[key2][7] = cc_Dict[key2][7] + row['TOTAL FEES']
                                cc_Dict[key2][8] = cc_Dict[key2][8] + row['PTO2']
                                cc_Dict[key2][9] = cc_Dict[key2][9] + row['Electronics Nontaxable']
                                cc_Dict[key2][10] = cc_Dict[key2][10] + row['Reimbursement-Non Taxable']
                                cc_Dict[key2][11] = cc_Dict[key2][11] + row['Total Client Charges']
                                cc_Dict[key2][12] = row['DEPT CODE']
                                cc_Dict[key2][13] = row['Employee Name']
                                cc_Dict[key2][14] = row['Pay End Date']
                                cc_Dict[key2][15] = row['Invoice Date']
                                cc_Dict[key2][16] = True
                                df_spring = df_spring.drop(index)

                        for key3 in rhq_Dict:
                            if (row['Employee Name'] == key3):
                                #print(key2, " ", i, " ", j, " ", k)
                                rhq_Dict[key3][0] =  rhq_Dict[key3][0] + row['Gross Wages']
                                rhq_Dict[key3][1] =  rhq_Dict[key3][1] + row['OT']
                                rhq_Dict[key3][2] =  rhq_Dict[key3][2] + row['Bonus']
                                rhq_Dict[key3][3] =  rhq_Dict[key3][3] + row['Taxes - ER - Totals']
                                rhq_Dict[key3][4] =  rhq_Dict[key3][4] + row['Workers Comp Fee - Totals']
                                rhq_Dict[key3][5] =  rhq_Dict[key3][5] + row['401k/Roth-ER']
                                rhq_Dict[key3][6] =  rhq_Dict[key3][6] + row['BENEFITS wo 401K']
                                rhq_Dict[key3][7] =  rhq_Dict[key3][7] + row['TOTAL FEES']
                                rhq_Dict[key3][8] =  rhq_Dict[key3][8] + row['PTO2']
                                rhq_Dict[key3][9] =  rhq_Dict[key3][9] + row['Electronics Nontaxable']
                                rhq_Dict[key3][10] =  rhq_Dict[key3][10] + row['Reimbursement-Non Taxable']
                                rhq_Dict[key3][11] =  rhq_Dict[key3][11] + row['Total Client Charges']
                                rhq_Dict[key3][12] = row['DEPT CODE']
                                rhq_Dict[key3][13] = row['Employee Name']
                                rhq_Dict[key3][14] = row['Pay End Date']
                                rhq_Dict[key3][15] = row['Invoice Date']
                                rhq_Dict[key3][16] = True
                                df_spring = df_spring.drop(index)

                        for key4 in mr_Dict:
                            if (row['Employee Name'] == key4):
                                #print(key2, " ", i, " ", j, " ", k)
                                mr_Dict[key4][0] =  mr_Dict[key4][0] + row['Gross Wages']
                                mr_Dict[key4][1] =  mr_Dict[key4][1] + row['OT']
                                mr_Dict[key4][2] =  mr_Dict[key4][2] + row['Bonus']
                                mr_Dict[key4][3] =  mr_Dict[key4][3] + row['Taxes - ER - Totals']
                                mr_Dict[key4][4] =  mr_Dict[key4][4] + row['Workers Comp Fee - Totals']
                                mr_Dict[key4][5] =  mr_Dict[key4][5] + row['401k/Roth-ER']
                                mr_Dict[key4][6] =  mr_Dict[key4][6] + row['BENEFITS wo 401K']
                                mr_Dict[key4][7] =  mr_Dict[key4][7] + row['TOTAL FEES']
                                mr_Dict[key4][8] =  mr_Dict[key4][8] + row['PTO2']
                                mr_Dict[key4][9] =  mr_Dict[key4][9] + row['Electronics Nontaxable']
                                mr_Dict[key4][10] =  mr_Dict[key4][10] + row['Reimbursement-Non Taxable']
                                mr_Dict[key4][11] =  mr_Dict[key4][11] + row['Total Client Charges']
                                mr_Dict[key4][12] = row['DEPT CODE']
                                mr_Dict[key4][13] = row['Employee Name']
                                mr_Dict[key4][14] = row['Pay End Date']
                                mr_Dict[key4][15] = row['Invoice Date']
                                mr_Dict[key4][16] = True
                                df_spring = df_spring.drop(index)

                        for key5 in fc_Dict:
                            if (row['Employee Name'] == key5):
                                #print(key2, " ", i, " ", j, " ", k)
                                fc_Dict[key5][0] =  fc_Dict[key5][0] + row['Gross Wages']
                                fc_Dict[key5][1] =  fc_Dict[key5][1] + row['OT']
                                fc_Dict[key5][2] =  fc_Dict[key5][2] + row['Bonus']
                                fc_Dict[key5][3] =  fc_Dict[key5][3] + row['Taxes - ER - Totals']
                                fc_Dict[key5][4] =  fc_Dict[key5][4] + row['Workers Comp Fee - Totals']
                                fc_Dict[key5][5] =  fc_Dict[key5][5] + row['401k/Roth-ER']
                                fc_Dict[key5][6] =  fc_Dict[key5][6] + row['BENEFITS wo 401K']
                                fc_Dict[key5][7] =  fc_Dict[key5][7] + row['TOTAL FEES']
                                fc_Dict[key5][8] =  fc_Dict[key5][8] + row['PTO2']
                                fc_Dict[key5][9] =  fc_Dict[key5][9] + row['Electronics Nontaxable']
                                fc_Dict[key5][10] =  fc_Dict[key5][10] + row['Reimbursement-Non Taxable']
                                fc_Dict[key5][11] =  fc_Dict[key5][11] + row['Total Client Charges']
                                fc_Dict[key5][12] = row['DEPT CODE']
                                fc_Dict[key5][13] = row['Employee Name']
                                fc_Dict[key5][14] = row['Pay End Date']
                                fc_Dict[key5][15] = row['Invoice Date']
                                fc_Dict[key5][16] = True
                                df_spring = df_spring.drop(index)

                        for key6 in co_Dict:
                            if (row['Employee Name'] == key6):
                                #print(key2, " ", i, " ", j, " ", k)
                                co_Dict[key6][0] =  co_Dict[key6][0] + row['Gross Wages']
                                co_Dict[key6][1] =  co_Dict[key6][1] + row['OT']
                                co_Dict[key6][2] =  co_Dict[key6][2] + row['Bonus']
                                co_Dict[key6][3] =  co_Dict[key6][3] + row['Taxes - ER - Totals']
                                co_Dict[key6][4] =  co_Dict[key6][4] + row['Workers Comp Fee - Totals']
                                co_Dict[key6][5] =  co_Dict[key6][5] + row['401k/Roth-ER']
                                co_Dict[key6][6] =  co_Dict[key6][6] + row['BENEFITS wo 401K']
                                co_Dict[key6][7] =  co_Dict[key6][7] + row['TOTAL FEES']
                                co_Dict[key6][8] =  co_Dict[key6][8] + row['PTO2']
                                co_Dict[key6][9] =  co_Dict[key6][9] + row['Electronics Nontaxable']
                                co_Dict[key6][10] =  co_Dict[key6][10] + row['Reimbursement-Non Taxable']
                                co_Dict[key6][11] =  co_Dict[key6][11] + row['Total Client Charges']
                                co_Dict[key6][12] = row['DEPT CODE']
                                co_Dict[key6][13] = row['Employee Name']
                                co_Dict[key6][14] = row['Pay End Date']
                                co_Dict[key6][15] = row['Invoice Date']
                                co_Dict[key6][16] = True
                                df_spring = df_spring.drop(index)




                for emp in exc_Dict:
                    if re.match('Krall,Audrey*', str(emp), re.IGNORECASE) or re.match('Dean,Ursula*', str(emp), re.IGNORECASE):
                    #if re.match('2495811*', str(emp)) or re.match('1906920*', str(emp)):
                        if (exc_Dict[emp][16] == True):
                        #  Loop Through Audrey Krall Locations in Exception List
                            for loc in SFOAKSV:
                                if loc == 'SF':
                                    pct = 0.5625
                                if loc == 'OAK':
                                    pct = 0.25
                                if loc == 'SV':
                                    pct = 0.1875
                                # Calculate Allocation Values
                                alloc_GrossWages_Sum = exc_Dict[emp][0] * pct
                                alloc_OT_Sum = exc_Dict[emp][1] * pct
                                alloc_Bonus_Sum = exc_Dict[emp][2] * pct
                                alloc_TaxesERTotals_Sum = exc_Dict[emp][3] * pct
                                alloc_WorkersCompFeeTot_Sum = exc_Dict[emp][4] * pct
                                alloc_Roth401kCombo_Sum = exc_Dict[emp][5] * pct
                                alloc_BenWO401k_Sum = exc_Dict[emp][6] * pct
                                alloc_TotalFees_Sum = exc_Dict[emp][7] * pct
                                alloc_PTO2_Sum = exc_Dict[emp][8] * pct
                                alloc_ElecNonTax_Sum = exc_Dict[emp][9] * pct
                                alloc_ReimbNonTax_Sum = exc_Dict[emp][10] * pct
                                alloc_TotClientCharges_Sum = exc_Dict[emp][11] * pct
                                empN = exc_Dict[emp][13]
                                deptCode = exc_Dict[emp][12]
                                pedExc = exc_Dict[emp][14]
                                ivdExc = exc_Dict[emp][15]
                                # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                                df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                    
                    if re.match('Tran,Nam D*', str(emp), re.IGNORECASE) and (exc_Dict[emp][16] == True):
                    #if re.match('1804091*', str(emp)) and (exc_Dict[emp][16] == True):                        
                        for loc in SFHQSV:
                            if loc == 'SF':
                                pct = 0.59
                            if loc == 'HQ':
                                pct = 0.1
                            if loc == 'SV':
                                pct = 0.31
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = exc_Dict[emp][0] * pct
                            alloc_OT_Sum = exc_Dict[emp][1] * pct
                            alloc_Bonus_Sum = exc_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = exc_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = exc_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = exc_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = exc_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = exc_Dict[emp][7] * pct
                            alloc_PTO2_Sum = exc_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = exc_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = exc_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = exc_Dict[emp][11] * pct
                            empN = exc_Dict[emp][13]
                            deptCode = exc_Dict[emp][12]
                            pedExc = exc_Dict[emp][14]
                            ivdExc = exc_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                    
                    if re.match('Klatsky,Peter*', str(emp), re.IGNORECASE) and (exc_Dict[emp][16] == True):
                    #if re.match('1770296*', str(emp)) and (exc_Dict[emp][16] == True):
                    
                        for loc in NYCMSO:
                            if loc == 'NYC':
                                pct = 0.75
                            else:
                                pct = 0.25
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = exc_Dict[emp][0] * pct
                            alloc_OT_Sum = exc_Dict[emp][1] * pct
                            alloc_Bonus_Sum = exc_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = exc_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = exc_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = exc_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = exc_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = exc_Dict[emp][7] * pct
                            alloc_PTO2_Sum = exc_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = exc_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = exc_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = exc_Dict[emp][11] * pct
                            empN = exc_Dict[emp][13]
                            deptCode = exc_Dict[emp][12]
                            pedExc = exc_Dict[emp][14]
                            ivdExc = exc_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                    
                    if re.match('Vaccari,Sergio*', str(emp), re.IGNORECASE) and (exc_Dict[emp][16] == True):
                    #if re.match('1785954*', str(emp)) and (exc_Dict[emp][16] == True):
                        #  Loop Through Sergio Vaccari Locations in Exception List
                        for loc in SFOAKSV:
                            if loc == 'SF':
                                pct = 0.5625
                            if loc == 'OAK':
                                pct = 0.25
                            if loc == 'SV':    
                                pct = 0.1875
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = exc_Dict[emp][0] * pct
                            alloc_OT_Sum = exc_Dict[emp][1] * pct
                            alloc_Bonus_Sum = exc_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = exc_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = exc_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = exc_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = exc_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = exc_Dict[emp][7] * pct
                            alloc_PTO2_Sum = exc_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = exc_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = exc_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = exc_Dict[emp][11] * pct
                            empN = exc_Dict[emp][13]
                            deptCode = exc_Dict[emp][12]
                            pedExc = exc_Dict[emp][14]
                            ivdExc = exc_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]

                    if re.match('Spivey,Allison Michele*', str(emp), re.IGNORECASE) and (exc_Dict[emp][16] == True):
                    #if re.match('1785954*', str(emp)) and (exc_Dict[emp][16] == True):
                        #  Loop Through Sergio Vaccari Locations in Exception List
                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':    
                                pct = 0.15
                            if loc == 'NYC':    
                                pct = 0.2
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = exc_Dict[emp][0] * pct
                            alloc_OT_Sum = exc_Dict[emp][1] * pct
                            alloc_Bonus_Sum = exc_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = exc_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = exc_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = exc_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = exc_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = exc_Dict[emp][7] * pct
                            alloc_PTO2_Sum = exc_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = exc_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = exc_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = exc_Dict[emp][11] * pct
                            empN = exc_Dict[emp][13]
                            deptCode = exc_Dict[emp][12]
                            pedExc = exc_Dict[emp][14]
                            ivdExc = exc_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                    
                    
                
                for emp in cc_Dict:
                    #  
                    if cc_Dict[emp][16] == True:

                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':    
                                pct = 0.15
                            if loc == 'NYC':
                                pct = 0.2
                            #if loc == 'SF':
                            #    pct = 0.34
                            #else:
                            #    pct = 0.33
                            
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = cc_Dict[emp][0] * pct
                            alloc_OT_Sum = cc_Dict[emp][1] * pct
                            alloc_Bonus_Sum = cc_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = cc_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = cc_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = cc_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = cc_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = cc_Dict[emp][7] * pct
                            alloc_PTO2_Sum = cc_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = cc_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = cc_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = cc_Dict[emp][11] * pct
                            empN = cc_Dict[emp][13]
                            deptCode = cc_Dict[emp][12]
                            pedExc = cc_Dict[emp][14]
                            ivdExc = cc_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                        
                for emp in mr_Dict:
                    #  
                    if mr_Dict[emp][16] == True:

                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':    
                                pct = 0.15
                            if loc == 'NYC':
                                pct = 0.2
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = mr_Dict[emp][0] * pct
                            alloc_OT_Sum = mr_Dict[emp][1] * pct
                            alloc_Bonus_Sum = mr_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = mr_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = mr_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = mr_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = mr_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = mr_Dict[emp][7] * pct
                            alloc_PTO2_Sum = mr_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = mr_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = mr_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = mr_Dict[emp][11] * pct
                            empN = mr_Dict[emp][13]
                            deptCode = mr_Dict[emp][12]
                            pedExc = mr_Dict[emp][14]
                            ivdExc = mr_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]

                for emp in rhq_Dict:
                    #  
                    if rhq_Dict[emp][16] == True:

                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':
                                pct = 0.15
                            if loc == 'NYC':
                                pct = 0.2
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = rhq_Dict[emp][0] * pct
                            alloc_OT_Sum = rhq_Dict[emp][1] * pct
                            alloc_Bonus_Sum = rhq_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = rhq_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = rhq_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = rhq_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = rhq_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = rhq_Dict[emp][7] * pct
                            alloc_PTO2_Sum = rhq_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = rhq_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = rhq_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = rhq_Dict[emp][11] * pct
                            empN = rhq_Dict[emp][13]
                            deptCode = rhq_Dict[emp][12]
                            pedExc = rhq_Dict[emp][14]
                            ivdExc = rhq_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]


                for emp in fc_Dict:
                    #  
                    if fc_Dict[emp][16] == True:

                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':
                                pct = 0.15
                            if loc == 'NYC':
                                pct = 0.2
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = fc_Dict[emp][0] * pct
                            alloc_OT_Sum = fc_Dict[emp][1] * pct
                            alloc_Bonus_Sum = fc_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = fc_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = fc_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = fc_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = fc_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = fc_Dict[emp][7] * pct
                            alloc_PTO2_Sum = fc_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = fc_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = fc_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = fc_Dict[emp][11] * pct
                            empN = fc_Dict[emp][13]
                            deptCode = fc_Dict[emp][12]
                            pedExc = fc_Dict[emp][14]
                            ivdExc = fc_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]

                for emp in co_Dict:
                    #
                    if re.match('Lee,My Dung*', str(emp), re.IGNORECASE) and (co_Dict[emp][16] == True):
                        #  Loop Through Locations in Exception List and calculate allocations.
                        for loc in SFOAKSVNYCNEST:
                            if loc == 'Nest':
                                pct = 0.1
                            if loc == 'SF':
                                pct = 0.405
                            if loc == 'OAK':
                                pct = 0.18
                            if loc == 'SV':
                                pct = 0.135
                            if loc == 'NYC':
                                pct = 0.18
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = co_Dict[emp][0] * pct
                            alloc_OT_Sum = co_Dict[emp][1] * pct
                            alloc_Bonus_Sum = co_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = co_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = co_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = co_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = co_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = co_Dict[emp][7] * pct
                            alloc_PTO2_Sum = co_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = co_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = co_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = co_Dict[emp][11] * pct
                            empN = co_Dict[emp][13]
                            deptCode = co_Dict[emp][12]
                            pedExc = co_Dict[emp][14]
                            ivdExc = co_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                    

                    elif re.match('Lagano,Lauren*', str(emp), re.IGNORECASE) and (co_Dict[emp][16] == True):
                    #if re.match('10328183*', str(emp)) and (co_Dict[emp][16] == True):
                        #  Loop Through Sergio Vaccari Locations in Exception List
                        for loc in HQNEST:
                            pct = 0.5
                            # Calculate Allocation Values
                            alloc_GrossWages_Sum = co_Dict[emp][0] * pct
                            alloc_OT_Sum = co_Dict[emp][1] * pct
                            alloc_Bonus_Sum = co_Dict[emp][2] * pct
                            alloc_TaxesERTotals_Sum = co_Dict[emp][3] * pct
                            alloc_WorkersCompFeeTot_Sum = co_Dict[emp][4] * pct
                            alloc_Roth401kCombo_Sum = co_Dict[emp][5] * pct
                            alloc_BenWO401k_Sum = co_Dict[emp][6] * pct
                            alloc_TotalFees_Sum = co_Dict[emp][7] * pct
                            alloc_PTO2_Sum = co_Dict[emp][8] * pct
                            alloc_ElecNonTax_Sum = co_Dict[emp][9] * pct
                            alloc_ReimbNonTax_Sum = co_Dict[emp][10] * pct
                            alloc_TotClientCharges_Sum = co_Dict[emp][11] * pct
                            empN = co_Dict[emp][13]
                            deptCode = co_Dict[emp][12]
                            pedExc = co_Dict[emp][14]
                            ivdExc = co_Dict[emp][15]
                            # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                            df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                    


                    else:
                        if co_Dict[emp][16] == True:

                            for loc in SFOAKSVNYC:
                                if loc == 'SF':
                                    pct = 0.45
                                if loc == 'OAK':
                                    pct = 0.2
                                if loc == 'SV':
                                    pct = 0.15
                                if loc == 'NYC':
                                    pct = 0.2
                                # Calculate Allocation Values
                                alloc_GrossWages_Sum = co_Dict[emp][0] * pct
                                alloc_OT_Sum = co_Dict[emp][1] * pct
                                alloc_Bonus_Sum = co_Dict[emp][2] * pct
                                alloc_TaxesERTotals_Sum = co_Dict[emp][3] * pct
                                alloc_WorkersCompFeeTot_Sum = co_Dict[emp][4] * pct
                                alloc_Roth401kCombo_Sum = co_Dict[emp][5] * pct
                                alloc_BenWO401k_Sum = co_Dict[emp][6] * pct
                                alloc_TotalFees_Sum = co_Dict[emp][7] * pct
                                alloc_PTO2_Sum = co_Dict[emp][8] * pct
                                alloc_ElecNonTax_Sum = co_Dict[emp][9] * pct
                                alloc_ReimbNonTax_Sum = co_Dict[emp][10] * pct
                                alloc_TotClientCharges_Sum = co_Dict[emp][11] * pct
                                empN = co_Dict[emp][13]
                                deptCode = co_Dict[emp][12]
                                pedExc = co_Dict[emp][14]
                                ivdExc = co_Dict[emp][15]
                                # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                                df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                                                    
                                            
    

   
    df_concatenated = pd.concat([df_spring, df_exceptions], ignore_index=True).fillna(0)
    
    df_concatenated.reset_index()
    # df_concatenated.to_excel('test.xlsx', index = False)
    df_group = df_concatenated.groupby(['Invoice Number', 'Department Long Descr', 'SUB_DEPARTMENT', 'LOCATION'])
    type(df_group)
    df_group.ngroups
    df_group.size()
    df_group.groups
    

    # Chart of Accounts

        # Create a dictionary representing the Chart of Accounts
    CoA = {4 : [61110, 51112, 51111, 51110, 61110, 61110, 51113], \
        5 : [61120, 51122, 51121, 51120, 61120, 61120, 51123], \
        6 : [23500, 23500, 23500, 23500, 23500, 23500, 23500], \
        7 : [61140, 51142, 51141, 51140, 61140, 61140, 51143], \
        8 : [61160, 51152, 51151, 51150, 61160, 61160, 51153], \
        9 : [61170, 51162, 51161, 51160, 61170, 61170, 51163], \
        10 : [61180, 51172, 51171, 51170, 61180, 61180, 51173], \
        11 : [69130, 51182, 51181, 51180, 69130, 69130, 51183], \
        12 : [23400, 23400, 23400, 23400, 23400, 23400, 23400], \
        13 : [65190, 65190, 65190, 65190, 65190, 65190, 65190], \
        14 : [22500, 22500, 22500, 22500, 22500, 22500, 22500]}

    # Create new Dataframe for the Output.
    df_Output = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', 'DebitAmt', 'CreditAmt', 'Loc', 'Dept', 'Provider', 'Service Line', 'Comments'])
    #print(df_Output)
    CoA_Index = 0
    for name, r in df_group:
        #print(name[2])
        if name[2] == 'HQ':
            CoA_Index = 0
        elif name[2] == 'Lab':
            CoA_Index = 1
        elif name[2] == 'ASC':
            CoA_Index = 2
        elif name[2] == 'Clinical':
            CoA_Index = 3
        elif name[2] == 'Operating':
            CoA_Index = 4
        elif name[2] == 'NEST':
            CoA_Index = 5
        elif name[2] == 'MD':
            CoA_Index = 6
        #print(CoA_Index)
        a = name
        b = r['Gross Wages'].sum()
        c = r['OT'].sum()
        d = r['Bonus'].sum()
        e = r['Taxes - ER - Totals'].sum()
        f = r['Workers Comp Fee - Totals'].sum()
        g = r['401k/Roth-ER'].sum()
        h = r['BENEFITS wo 401K'].sum()
        i = r['TOTAL FEES'].sum()
        j = r['PTO2'].sum()
        k = r['Electronics Nontaxable'].sum()
        l = r['Reimbursement-Non Taxable'].sum()
        m = r['Total Client Charges'].sum()

        """
        ## Troubleshooting Code ##
        bigsum = b+c+d+e+f+g+h+i+j+k+l
            if abs(bigsum - m) > 1:
            print(r)
            print(a, bigsum, m)
        """

        ped = r['Pay End Date']
        # Convert data type to datetime64[ns]
        ped = ped.astype("datetime64[ns]")
        ped_s = str(ped).split('Name', 1)[0]
        ped_s = ped_s[len(ped_s) - 11:]
        ivd = r['Invoice Date']
        # Convert data type to datetime64[ns]
        ivd = ivd.astype("datetime64[ns]")
        ivd_s = str(ivd).split('Name', 1)[0]
        ivd_s = ivd_s[len(ivd_s) - 11:]
        #ivd_s = re.sub(r"[a-zA-Z]","",str(ivd_s)[3:10])
        deptCode = r['DEPT CODE']
        dp = re.sub(r"[^0-9]","",str(deptCode)[3:10])

        
        if m != 0:
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[4][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), b, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[5][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), c, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[6][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), d, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[7][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), e, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[8][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), f, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[9][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), g, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[10][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), h, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[11][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), i, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[12][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), j, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[13][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), k, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], CoA[14][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), l, "", name[3], dp, "", "", ""]
            df_Output.loc[len(df_Output.index)] = ["", ped_s, ivd_s, "", name[2], 23300, "", str(name[0]) + ' ' + str(name[1]), "", m, name[3], dp, "", "", ""]
        

    
    inp = input("Please type name of file for Output:")
    des = str(inp + '.xlsx')
    df_Output.to_excel(des, index = False)

    runningtime = time.time() - start
    print("The time for this script is:", runningtime)

    
    

def FilePrompt():
    root = tk.Tk()
    root.title('Tkinter Open File Dialog')
    root.resizable(False, False)
    root.geometry('300x150')
    root.withdraw()


    filename = fd.askopenfilename()

    return filename
    


if __name__ == "__main__":
    main()
