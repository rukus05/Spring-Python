

from msilib.schema import File
import time
import pandas as pd
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
    mr_Dict = {}    # Call Center Dict
    AK = ['SF', 'OAK', 'SV']
    QTR = ['SF', 'OAK', 'SV', 'NYC']
    MDL = ['SF', 'OAK', 'SV', 'NYC', 'Nest']
    LL = ['HQ', 'Nest']
    dldloc = df_spring.columns.get_loc('Department Long Descr')
    locloc = df_spring.columns.get_loc('LOCATION')
    # First group of 4 For loops is to Handle (Clean) Exceptions
    # By using 4 nested FOR loops, we can group the rows by Invoices, Dept Long Desc, Sup Dept, and Location
    
    # Reclassify some exceptions before massaging the dataframe
    for index, row in df_spring.iterrows():
        if (row['Employee Name'] == 'Cicciarello,Claire') or (row['Employee Name'] == 'Mock,Gina M'):
            df_spring.at[index, 'Department Long Descr'] = 'Operating'
            df_spring.at[index, 'LOCATION'] = 'NYC'
        if row['Employee Name'] == 'Lee,Stephannie Victoria':
            df_spring.at[index, 'Department Long Descr'] = 'Call Center'
            df_spring.at[index, 'LOCATION'] = 'HQ'
    # Add all Call Center People (Not Stephannie Lee) into a Dict Data Structure
    # Add all Medical Records People into a Dict Data Structure
    for index, row in df_spring.iterrows():
        if (row['Department Long Descr'] == 'Call Center') and (row['Employee Name'] != 'Lee,Stephannie Victoria'): 
            cc_Dict[row['Employee Name']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
        if (row['Department Long Descr'] == 'Medical Records'): 
            mr_Dict[row['Employee Name']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
    
    
    for i in uniqueInvoices:
        for k in uniqueSub_Dept:
            for j in unique_DLD:
                
                # Handle Exceptions
                exc_Dict["Krall,Audrey"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Dam,Phuong My"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Lee,My Dung"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Trieu,Minh Hue"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict[ "Lagano,Lauren"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Bell,Allie Marie"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                for key in cc_Dict:
                    cc_Dict[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                for key in mr_Dict:
                    mr_Dict[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]

                for index, row in df_spring.iterrows():
                    if (row['Invoice Number'] == i) and (row['Department Long Descr'] == j) and (row['SUB_DEPARTMENT'] == k):
                        ## Try using this here:  
                        ## if row[0] in cs_revenue_dict:
                        for key in exc_Dict:
                            if (row['Employee Name'] == key):
                                #print(key, " ", i, " ", j, " ", k)
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

                        for key3 in mr_Dict:
                            if (row['Employee Name'] == key3):
                                #print(key2, " ", i, " ", j, " ", k)
                                mr_Dict[key3][0] =  mr_Dict[key3][0] + row['Gross Wages']
                                mr_Dict[key3][1] =  mr_Dict[key3][1] + row['OT']
                                mr_Dict[key3][2] =  mr_Dict[key3][2] + row['Bonus']
                                mr_Dict[key3][3] =  mr_Dict[key3][3] + row['Taxes - ER - Totals']
                                mr_Dict[key3][4] =  mr_Dict[key3][4] + row['Workers Comp Fee - Totals']
                                mr_Dict[key3][5] =  mr_Dict[key3][5] + row['401k/Roth-ER']
                                mr_Dict[key3][6] =  mr_Dict[key3][6] + row['BENEFITS wo 401K']
                                mr_Dict[key3][7] =  mr_Dict[key3][7] + row['TOTAL FEES']
                                mr_Dict[key3][8] =  mr_Dict[key3][8] + row['PTO2']
                                mr_Dict[key3][9] =  mr_Dict[key3][9] + row['Electronics Nontaxable']
                                mr_Dict[key3][10] =  mr_Dict[key3][10] + row['Reimbursement-Non Taxable']
                                mr_Dict[key3][11] =  mr_Dict[key3][11] + row['Total Client Charges']
                                mr_Dict[key3][12] = row['DEPT CODE']
                                mr_Dict[key3][13] = row['Employee Name']
                                mr_Dict[key3][14] = row['Pay End Date']
                                mr_Dict[key3][15] = row['Invoice Date']
                                mr_Dict[key3][16] = True
                                df_spring = df_spring.drop(index)


                for emp in exc_Dict:
                    
                    if (emp == 'Krall,Audrey') and (exc_Dict[emp][16] == True):
                        #  Loop Through Audrey Krall Locations in Exception List
                        for loc in AK:
                            if loc == 'SF':
                                pct = 0.70
                            else:
                                pct = 0.15
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
                    
                    if (emp == 'Dam,Phuong My') or (emp == 'Trieu,Minh Hue') or (emp == 'Bell,Allie Marie'):
                        if exc_Dict[emp][16] == True:
                        #  Loop Through Audrey Krall Locations in Exception List
                            for loc in QTR:
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
                                df_exceptions.loc[len(df_exceptions.index)] = [empN, i, pedExc, ivdExc, loc, k, j, deptCode,  alloc_GrossWages_Sum, alloc_OT_Sum, alloc_Bonus_Sum, alloc_TaxesERTotals_Sum, alloc_WorkersCompFeeTot_Sum, alloc_Roth401kCombo_Sum, alloc_BenWO401k_Sum,  alloc_TotalFees_Sum, alloc_PTO2_Sum, alloc_ElecNonTax_Sum, alloc_ReimbNonTax_Sum, alloc_TotClientCharges_Sum]
                    
                    if (emp == 'Lee,My Dung') and (exc_Dict[emp][16] == True):
                        #  Loop Through Audrey Krall Locations in Exception List
                        for loc in MDL:
                            if loc == 'Nest':
                                pct = 0.1
                            else:
                                pct = 0.225
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
                    
                    if (emp == 'Vaccari,Sergio') and (exc_Dict[emp][16] == True):
                        #  Loop Through Sergio Vaccari Locations in Exception List
                        for loc in AK:
                            if loc == 'SF':
                                pct = 0.34
                            else:
                                pct = 0.33
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
                    
                    if (emp == 'Lagano,Lauren') and exc_Dict[emp][16] == True:
                        #  Loop Through Sergio Vaccari Locations in Exception List
                        for loc in LL:
                            pct = 0.5
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

                        for loc in AK:
                            if loc == 'SF':
                                pct = 0.34
                            else:
                                pct = 0.33
                            
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

                        for loc in QTR:
                            pct = 0.25
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
    print(cc_Dict)
    print(mr_Dict)
    print(exc_Dict)
                        
 

   
    df_concatenated = pd.concat([df_spring, df_exceptions], ignore_index=True).fillna(0)
    
    df_concatenated.reset_index()

    inp = input("Please type name of file for Output:")
    des = str(inp + '.xlsx')
    df_concatenated.to_excel(des, index = False)

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



