

from msilib.schema import File
import pandas as pd
import openpyxl
import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.filedialog import asksaveasfile


def main(): 
    
    # For this program to work properly, the CoA_Pandas.py file needs to be in the same directory as the RawData.xlsx file.
    # Read in Data from the "RawData.xlsx" file.

    f = FilePrompt()
    df_spring = pd.read_excel(f)

    #print(datetime.datetime.now().strftime('%Y-%m-%d'))
   
    """
    # The below code iterates over a row and gets a particular column's values.

    df_spring = df_spring.reset_index()
    for index, row in df_spring.iterrows():
        print(row['LOCATION'], row['SUB_DEPARTMENT'], row['Gross Wages - Totals'])

    """
    df_spring = df_spring.reset_index()
    
    # Get the unique Invoice Numbers, Locations, Sub Departments, and Department Long Descriptions.  This is needed to loop through and group them properly.
    uniqueInvoices = df_spring['Invoice Number'].unique()
    uniqueLocations = df_spring['LOCATION'].unique()
    uniqueSub_Dept = df_spring['SUB_DEPARTMENT'].unique()
    unique_DLD = df_spring['Department Long Descr'].unique()
    #print(unique_DLD)
    #print(df_spring.info())
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
    #df_spring['Pay End Date'] = df_spring['Pay End Date']
    
    #df_forxceptions=df_spring.copy()
    
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
        14 : [65190, 65190, 65190, 65190, 65190, 65190, 65190]} 
    """
    # TEST CODE
    unique_DLD = df_spring['Department Long Descr']   
    This prints an entire Column:
    print(uniqueLocations)
    print(uniqueSub_Dept)
    print(uniqueInvoices)
    GrossWages_Total = 0
    print(unique_DLD)
    """
    # Create new Dataframe for the Exceptions Output.
    df_exceptions = pd.DataFrame(columns=['Employee Name', 'Invoice Number', 'Pay End Date', 'Invoice Date', 'LOCATION', 'SUB_DEPARTMENT', 'Department Long Descr', 'DEPT CODE', 'Gross Wages', 'OT', 'Bonus', 'Taxes - ER - Totals', 'Workers Comp Fee - Totals', '401k/Roth-ER', 'BENEFITS wo 401K', 'TOTAL FEES', 'PTO2', 'Electronics Nontaxable', 'Reimbursement-Non Taxable', 'Total Client Charges'])
    
    # Create new Dataframe for the Output.
    df_Output = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', 'DebitAmt', 'CreditAmt', 'Loc', 'Dept', 'Provider', 'Service Line', 'Comments'])
    # print(df_Output)
    # Exceptions List (Key)
    # 
    # alloc_dict = {'Krall,Audrey' : []}
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
    
    
    #print(cc_Dict.keys())
    #df_spring.to_excel('alloc_test output.xlsx', index = False)
    
    for i in uniqueInvoices:
        for k in uniqueSub_Dept:
            for j in unique_DLD:
                
                # Handle Exceptions
                exc_Dict["Krall,Audrey"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Dam,Phuong My"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Lee,My Dung"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict["Trieu,Minh Hue"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                exc_Dict[ "Lagano,Lauren"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
                for key in cc_Dict:
                    cc_Dict[key][16] = False
                for key in mr_Dict:
                    mr_Dict[key][16] = False

                for index, row in df_spring.iterrows():
                    if (row['Invoice Number'] == i) and (row['Department Long Descr'] == j) and (row['SUB_DEPARTMENT'] == k):
                                        
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

                #print(exc_Dict)
                #summary = [i, j, k , l, Exception_GrossWages_Sum, Exception_OT_Sum, Exception_Bonus_Sum, Exception_TaxesERTotals_Sum, Exception_WorkersCompFeeTot_Sum, Exception_Roth401kCombo_Sum, Exception_BenWO401k_Sum, Exception_TotalFees_Sum, Exception_PTO2_Sum, Exception_ElecNonTax_Sum, Exception_ReimbNonTax_Sum, Exception_TotClientCharges_Sum]          
                
                
                
                for emp in exc_Dict:
                    
                    if (emp == 'Krall,Audrey') and (exc_Dict[emp][16] == True):
                        #  Loop Through Audrey Krall Locations in Exception List
                        for loc in QTR:
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
                    
                    if (emp == 'Dam,Phuong My') or (emp == 'Trieu,Minh Hue'):
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
                            #print(emp," ", k, " ", j)
                            
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
                        
                        
    
    #magic = df_exceptions.reindex(df_spring)

    #print (df_exceptions)
    #df_exceptions.to_excel('alloc_test.xlsx', index = False)


   
    df_concatenated = pd.concat([df_spring, df_exceptions], ignore_index=True).fillna(0)
    
    df_concatenated.reset_index()
    des = SavePrompt()
    df_concatenated.to_excel(str(des), index = False)
    

    """
    # By using 4 nested FOR loops, we can group the rows by Invoices, Dept Long Desc, Sup Dept, and Location
    for i in uniqueInvoices:
        for j in unique_DLD:
            for k in uniqueSub_Dept:
                # Set the Index for the List (Array)  in the Chart of Accounts Dictionary, CoA
                if k == 'HQ':
                    CoA_Index = 0
                elif k == 'Lab':
                    CoA_Index = 1
                elif k == 'ASC':
                    CoA_Index = 2
                elif k == 'Clinical':
                    CoA_Index = 3
                elif k == 'Operating':
                    CoA_Index = 4
                elif k == 'NEST':
                    CoA_Index = 5
                elif k == 'MD':
                    CoA_Index = 6

                for l in uniqueLocations:
                    # Initialize Sum Variables
                    GrossWages_Sum = 0
                    OT_Sum = 0
                    Bonus_Sum = 0
                    TaxesERTotals_Sum = 0
                    WorkersCompFeeTot_Sum = 0
                    Roth401kCombo_Sum = 0
                    BenWO401k_Sum = 0
                    TotalFees_Sum = 0
                    PTO2_Sum = 0
                    ElecNonTax_Sum = 0
                    ReimbNonTax_Sum = 0
                    TotClientCharges_Sum = 0

                    # Loop through rows of the Raw Data file.
                    for index, row in df_concatenated.iterrows():
                        # These if statements force a match for specific Invoices, Dept Long Desc, Sup Dept, and Locations
                        if (row['Invoice Number'] == i) and (row['Department Long Descr'] == j) and (row['SUB_DEPARTMENT'] == k) and (row['LOCATION'] == l):
                            # Sum up the pertinent columns.
                            GrossWages_Sum = GrossWages_Sum + row['Gross Wages']
                            OT_Sum = OT_Sum + row['OT']
                            Bonus_Sum = Bonus_Sum + row['Bonus']
                            TaxesERTotals_Sum = TaxesERTotals_Sum + row['Taxes - ER - Totals']
                            WorkersCompFeeTot_Sum = WorkersCompFeeTot_Sum + row['Workers Comp Fee - Totals']
                            Roth401kCombo_Sum = Roth401kCombo_Sum + row['401k/Roth-ER']
                            BenWO401k_Sum = BenWO401k_Sum + row['BENEFITS wo 401K']
                            TotalFees_Sum = TotalFees_Sum + row['TOTAL FEES']
                            PTO2_Sum = PTO2_Sum + row['PTO2']
                            ElecNonTax_Sum = ElecNonTax_Sum + row['Electronics Nontaxable']
                            ReimbNonTax_Sum = ReimbNonTax_Sum + row['Reimbursement-Non Taxable']
                            TotClientCharges_Sum = TotClientCharges_Sum + row['Total Client Charges']
                            deptCode = row['DEPT CODE']
                            ped = row['Pay End Date']
                            ivd = row['Invoice Date']


                    # Create an array holding the Sums for a particular Invoice, DLD, Location, and Sub_Dept        
                    summary = [i, j, k , l, GrossWages_Sum, OT_Sum, Bonus_Sum, TaxesERTotals_Sum, WorkersCompFeeTot_Sum, Roth401kCombo_Sum, BenWO401k_Sum, TotalFees_Sum, PTO2_Sum, ElecNonTax_Sum, ReimbNonTax_Sum, TotClientCharges_Sum]          
                                    
                    
                    # Initialize counter for CoA
                    cnt = 4
                    # Loop over the Sum variables in the Summary Array.  
                    # Add row to output file if GrossWages doesn't equal 0.  
                    # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                    for x in summary[4:15]:
                        df_Output.loc[len(df_Output.index)] = ["", ped, ivd, "", str(k), CoA[cnt][CoA_Index], "", str(i) + ' ' + str(j), x, "", l, deptCode, "", "", ""]
                        cnt = cnt + 1       # Increment CoA counter
                            
                    # Add a credit entry to match the above debits.
                    df_Output.loc[len(df_Output.index)] = [ "", ped, ivd, "", str(k), 23300, "", str(i) + ' ' + str(j), "", TotClientCharges_Sum , l, deptCode, "", "", ""]
                    
                   
                    # Test Code
                    #print(row['Invoice Number'], row['Department Long Descr'], row['SUB_DEPARTMENT'], row['LOCATION'], GrossWages_Sum, Roth401kCombo_Sum)
    
    #df_concatenated = pd.concat([df_Output, df_exceptions])

    

    #  Export the Pandas DataFrame to an Excel file called "Output_test.xlsx"
    df_Output.to_excel('output_test.xlsx', index = False)
    """
def FilePrompt():
    root = tk.Tk()
    root.title('Tkinter Open File Dialog')
    root.resizable(False, False)
    root.geometry('300x150')
    root.withdraw()


    filename = fd.askopenfilename()

    return filename
    
def SavePrompt():
    root = tk.Tk()
    root.title('Tkinter Save File Dialog')
    root.resizable(False, False)
    root.geometry('300x150')
   

    files = [('All Files', '*.*'), 
             ('Excel Document', '*.xlsx')]
    g = asksaveasfile(filetypes = files, defaultextension = files)
  
    btn = ttk.Button(root, text = 'Save', command = lambda : save())
    btn.pack(side = TOP, pady = 20)

    return g

if __name__ == "__main__":
    main()



