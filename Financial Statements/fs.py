from math import nan
import time
import numpy as np
import xlsxwriter
from numpy import dtype, float64
import pandas as pd
import datetime
import tkinter as tk
from tkinter import TOP, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.filedialog import asksaveasfile
import PySimpleGUI as sg


def main(): 
    
    start = time.time()
    mos = int(input("How many months are in this report? "))
    
    
    
    #########
    # Delete this after troubleshooting
    #########
    #mos = 4
    #moslist = [0] * (mos + 1)
    #########

    # Initialize Dictionaries
    revenue_dict={}
    COGS_dict={}
    oe_dict={}
    nonOI_dict = {}
    cs_revenue_dict={}
    cs_COGS_dict={}
    cs_oe_dict={}
    cs_nonOI_dict = {}
    
    
    ##
    # Undo after troubleshooting
    f = FilePrompt()
    df_spring = pd.read_excel(f)
    ##

    #df_spring = pd.read_excel('RawData_FS.xlsx')

    
    
    df_spring = df_spring.reset_index()
    #  Remove spaces from column names so that itertuples() can be employed, vs iterrows()
    df_spring.columns = df_spring.columns.str.replace(' ','_')
    df_spring['Amount'] = df_spring['Amount'].fillna(0)
    
    uniquePL = df_spring['PL_Category'].unique()
    uniqueLoc = df_spring['Loc_Code_Dimension'].unique()
    
    
    df_spring['Posting_Date'] = pd.to_datetime(df_spring['Posting_Date'])
    df_spring['Posting_Date'] = df_spring['Posting_Date'].dt.month

    # Create Header for Columns when DF is exported (Held in cmos variable)
    cmos = ['Account']
    for x in range(mos):
        cmos.append('Month ' + str(x + 1))
    cmos.append('Total')
    
    


    for l in uniqueLoc:
        # Initialize Lists for Totals
        TR = []
        TCOGS = []
        GM = []
        TOE= []
        TNOE = []
        NI = []
        depreciation = [0] * (mos + 1)
        intexpense = [0] * (mos + 1)
        EBITDA = []
        # Create name for file to hold location DF export
        dfname = 'df_' + str(l)
        nm = dfname + '.xlsx'
        df_Output = pd.DataFrame(columns= cmos)
        # Initialize Dictionary values to 0.
        revenue_dict['Self-pay revenue'] = [0] * (mos + 1)
        revenue_dict['Commercial Insurance revenue'] = [0] * (mos + 1)
        revenue_dict['Progyny & Stork revenue'] = [0] * (mos + 1)
        revenue_dict['Storage revenue'] = [0] * (mos + 1)
        revenue_dict['Medication'] = [0] * (mos + 1)
        revenue_dict['Nest'] = [0] * (mos + 1)
        revenue_dict['Other Revenue'] = [0] * (mos + 1)

        COGS_dict['MD payroll'] = [0] * (mos + 1)
        COGS_dict['Clinical payroll'] = [0] * (mos + 1)
        COGS_dict['Lab payroll'] = [0] * (mos + 1)
        COGS_dict['ASC payroll'] = [0] * (mos + 1)
        COGS_dict['Supplies'] = [0] * (mos + 1)
        COGS_dict['Medication'] = [0] * (mos + 1)
        COGS_dict['Medical services'] = [0] * (mos + 1)
        
        oe_dict['Payroll'] = [0] * (mos + 1)
        oe_dict['Marketing'] = [0] * (mos + 1)
        oe_dict['Professional fees'] = [0] * (mos + 1)
        oe_dict['Rent'] = [0] * (mos + 1)
        oe_dict['Facilities'] = [0] * (mos + 1)
        oe_dict['Travel'] = [0] * (mos + 1)
        oe_dict['Facilities'] = [0] * (mos + 1)
        oe_dict['Employee related expenses'] = [0] * (mos + 1)
        oe_dict['Taxes & Regulatory'] = [0] * (mos + 1)
        oe_dict['Bank charges'] = [0] * (mos + 1)
        oe_dict['Other'] = [0] * (mos + 1)

        nonOI_dict['Non-operating income/(expense)'] = [0] * (mos + 1)

        # Add first item in Totals lists--this will be the first column in the totals row
        TR.append('Monthly Total Revenue')
        TCOGS.append('Monthly Total COGS')
        GM.append('Monthly GROSS MARGIN')
        TOE.append('Monthly Total OpEx')
        TNOE.append('Monthly Total Non OpEx')
        NI.append('Monthly Net Income')
        EBITDA.append('Monthly EBITA')
        # Initialize Depreciaton and Interest Expense Variables
        
        
        # Iterate through df_spring DF
        for row in df_spring.itertuples():
        # for index, row in df_spring.iterrows():
            x = str(row.PL_Category)
            m = int(row.Posting_Date)
            
            # Based on location, add the PL values.
            # A for loop for each section of the finanancial statement (dict)
            for key in revenue_dict:
                if (row.PL_Category == key) and (row.Loc_Code_Dimension == l):
                    r = row.Amount
                    revenue_dict[key][m-1] = revenue_dict[key][m-1] + r
            for key in COGS_dict:
                if (row.PL_Category == key) and (row.Loc_Code_Dimension == l):
                    r = row.Amount
                    COGS_dict[key][m-1] = COGS_dict[key][m-1] + r
            for key in oe_dict:
                if (row.PL_Category == key) and (row.Loc_Code_Dimension == l):
                    r = row.Amount
                    oe_dict[key][m-1] = oe_dict[key][m-1] + r

            for key in nonOI_dict:
                if (row.PL_Category == key) and (row.Loc_Code_Dimension == l):
                    r = row.Amount
                    nonOI_dict[key][m-1] = nonOI_dict[key][m-1] + r

            if (row.Loc_Code_Dimension == l) and (row.GL_Account == '71140Depreciation'):
                depreciation[m-1] = depreciation[m-1] + row.Amount
            
            if (row.Loc_Code_Dimension == l) and (row.GL_Account == '71130Interest expense'):
                intexpense[m-1] = intexpense[m-1] + row.Amount

        # Insert a blank row
        df_Output.loc[len(df_Output.index)] = nan
        # Add a row to indicate state of Revenue section of Financial statement
        df_Output.loc[len(df_Output.index)] = 'Revenue'
        # Function PLtoDF not working
        # df_Output = PLtoDF(revenue_dict, mos, df_Output)
        for key in revenue_dict:
            rowlist = [key]
            # Sum up the values of this key, and add to the last value in the dictionary list (Sum horizontially)
            tot = sum(revenue_dict[key])
            revenue_dict[key][mos] = tot
            #  Append all values of the dictionary key to the List and insert into the dataframe
            for m in revenue_dict[key]:
                rowlist.append(m)
            df_Output.loc[len(df_Output.index)] = rowlist
        # Calculate Total Revenue (Vertically) and add to DataFrame
        TR = MonthlyTotals(revenue_dict, mos, TR)
        df_Output.loc[len(df_Output.index)] = TR
        
        # Insert a blank row
        df_Output.loc[len(df_Output.index)] = nan

        # Add a row to indicate state of COGS section of Financial statement
        df_Output.loc[len(df_Output.index)] = 'COGS'
        
        for key in COGS_dict:
            rowlist = [key]
            tot = sum(COGS_dict[key])
            COGS_dict[key][mos] = tot
            for m in COGS_dict[key]:
                rowlist.append(m)
            df_Output.loc[len(df_Output.index)] = rowlist
        # Calculate Total COGS (Vertically) and add to DataFrame
        TCOGS = MonthlyTotals(COGS_dict, mos, TCOGS)
        df_Output.loc[len(df_Output.index)] = TCOGS

        # Insert a blank row
        df_Output.loc[len(df_Output.index)] = nan
        # Calculate Gross Margin and Add to DF
        # GM.append('GROSS MARGIN')
        for z in range(mos + 1):
            GM.append(TR[z + 1] - TCOGS[z + 1])
        df_Output.loc[len(df_Output.index)] = GM
        
        # Insert a blank row
        df_Output.loc[len(df_Output.index)] = nan
        
        df_Output.loc[len(df_Output.index)] = 'OpEx'

        for key in oe_dict:
            rowlist = [key]
            tot = sum(oe_dict[key])
            oe_dict[key][mos] = tot
            for m in oe_dict[key]:
                rowlist.append(m)
            df_Output.loc[len(df_Output.index)] = rowlist
        # Calculate Total OpEx (Vertically) and add to DataFrame
       
        TOE = MonthlyTotals(oe_dict, mos, TOE)
        df_Output.loc[len(df_Output.index)] = TOE
        
        # Insert a blank row
        df_Output.loc[len(df_Output.index)] = nan

        df_Output.loc[len(df_Output.index)] = 'Non OpEx'
        for key in nonOI_dict:
            rowlist = [key]
            tot = sum(nonOI_dict[key])
            nonOI_dict[key][mos] = tot
            for m in nonOI_dict[key]:
                rowlist.append(m)
            df_Output.loc[len(df_Output.index)] = rowlist
        # Calculate Total COGS (Vertically) and add to DataFrame
        
        TNOE = MonthlyTotals(nonOI_dict, mos, TNOE)
        df_Output.loc[len(df_Output.index)] = TNOE

        # Insert a blank row
        df_Output.loc[len(df_Output.index)] = nan

        # Calculate Net Income and Add to DF
        for z in range(mos + 1):
            NI.append(GM[z + 1] - TOE[z + 1] + TNOE[z + 1])
        df_Output.loc[len(df_Output.index)] = NI

        # Insert a blank row
        df_Output.loc[len(df_Output.index)] = nan

        # Calculate EBITDA and Add to DF
        # Insert here
        
        for z in range(mos + 1):
            EBITDA.append(NI[z + 1] + depreciation[z] + intexpense[z])
        df_Output.loc[len(df_Output.index)] = EBITDA
        

        # Copy the DataFrame to the appropriate site dataframe
        if l == 'DAN':
            DAN_df = df_Output.copy()
        elif l == 'HQ':
            HQ_df = df_Output.copy()
        elif l == 'nan':
            nan_df = df_Output.copy()
        elif l == 'NYC':
            NYC_df = df_Output.copy()
        elif l == 'OAK':
            OAK_df = df_Output.copy()
        elif l == 'RMT':
            RMT_df = df_Output.copy()
        elif l == 'RWC':
            RWC_df = df_Output.copy()
        elif l == 'SF':
            SF_df = df_Output.copy()
        elif l == 'SOMA':
            SOMA_df = df_Output.copy()
        elif l == 'SV':
            SV_df = df_Output.copy()
        elif l == 'VAN':
            VAN_df = df_Output.copy()
        
        

        #df_Output.to_excel(nm, index = False)       

    # Create Dataframe for the Consolidated Financial Statement
    df_consolidated = pd.concat([DAN_df, HQ_df, NYC_df, OAK_df, RMT_df, RWC_df, SF_df, SOMA_df, SV_df, VAN_df])

    # This section cleans the Consoldiated Statement DF

    # Set CS Dictionary values to 0
    cs_revenue_dict['Self-pay revenue'] = [0] * (mos + 1)
    cs_revenue_dict['Commercial Insurance revenue'] = [0] * (mos + 1)
    cs_revenue_dict['Progyny & Stork revenue'] = [0] * (mos + 1)
    cs_revenue_dict['Storage revenue'] = [0] * (mos + 1)
    cs_revenue_dict['Medication'] = [0] * (mos + 1)
    cs_revenue_dict['Nest'] = [0] * (mos + 1)
    cs_revenue_dict['Other Revenue'] = [0] * (mos + 1)
    cs_COGS_dict['MD payroll'] = [0] * (mos + 1)
    cs_COGS_dict['Clinical payroll'] = [0] * (mos + 1)
    cs_COGS_dict['Lab payroll'] = [0] * (mos + 1)
    cs_COGS_dict['ASC payroll'] = [0] * (mos + 1)
    cs_COGS_dict['Supplies'] = [0] * (mos + 1)
    cs_COGS_dict['Medication'] = [0] * (mos + 1)
    cs_COGS_dict['Medical services'] = [0] * (mos + 1)
    cs_oe_dict['Payroll'] = [0] * (mos + 1)
    cs_oe_dict['Marketing'] = [0] * (mos + 1)
    cs_oe_dict['Professional fees'] = [0] * (mos + 1)
    cs_oe_dict['Rent'] = [0] * (mos + 1)
    cs_oe_dict['Facilities'] = [0] * (mos + 1)
    cs_oe_dict['Travel'] = [0] * (mos + 1)
    cs_oe_dict['Facilities'] = [0] * (mos + 1)
    cs_oe_dict['Employee related expenses'] = [0] * (mos + 1)
    cs_oe_dict['Taxes & Regulatory'] = [0] * (mos + 1)
    cs_oe_dict['Bank charges'] = [0] * (mos + 1)
    cs_oe_dict['Other'] = [0] * (mos + 1)
    cs_nonOI_dict['Non-operating income/(expense)'] = [0] * (mos + 1)
    cs_EBITA = [0] * (mos + 1)

    for row in df_consolidated.itertuples(index = False):
        if row[0] in cs_revenue_dict:
            b = row[0]
            for a in range(mos + 1):
                cs_revenue_dict[b][a] = (cs_revenue_dict[b][a]) + (row[a + 1])
        elif row[0] in cs_COGS_dict:
            b = row[0]
            for a in range(mos + 1):
                cs_COGS_dict[b][a] = (cs_COGS_dict[b][a]) + (row[a + 1])
        elif row[0] in cs_oe_dict:
            b = row[0]
            for a in range(mos + 1):
                cs_oe_dict[b][a] = (cs_oe_dict[b][a]) + (row[a + 1])
        elif row[0] in cs_nonOI_dict:
            b = row[0]
            for a in range(mos + 1):
                cs_nonOI_dict[b][a] = (cs_nonOI_dict[b][a]) + (row[a + 1])

    #print(cs_revenue_dict)
    
    # Create new DF to hold cleaned dataframe
    df_cons_Output = pd.DataFrame(columns= cmos)
    
    c_TR = []
    c_TCOGS = []
    c_GM = []
    c_TOE= []
    c_TNOE = []
    c_NI = []
    c_EBITDA = []

    c_TR.append('Monthly Total Revenue')
    c_TCOGS.append('Monthly Total COGS')
    c_GM.append('Monthly GROSS MARGIN')
    c_TOE.append('Monthly Total OpEx')
    c_TNOE.append('Monthly Total Non OpEx')
    c_NI.append('Monthly Net Income')
    c_EBITDA.append('Monthly EBITA')
    # Add a row to indicate state of Revenue section of Financial statement
    df_cons_Output.loc[len(df_cons_Output.index)] = 'Revenue'

    # Place Consolidate Revenue Account info into DF
    for key2, values in cs_revenue_dict.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2
    
    # Calculate Total Revenue (Vertically) and add to DataFrame
    
    c_TR = MonthlyTotals(cs_revenue_dict, mos, c_TR)
    df_cons_Output.loc[len(df_cons_Output.index)] = c_TR

    # Insert a blank row
    df_cons_Output.loc[len(df_cons_Output.index)] = nan

    df_cons_Output.loc[len(df_cons_Output.index)] = 'COGS'
    for key2, values in cs_COGS_dict.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2

    # Calculate Total COGS (Vertically) and add to DataFrame
    
    c_TCOGS = MonthlyTotals(cs_COGS_dict, mos, c_TCOGS)
    df_cons_Output.loc[len(df_cons_Output.index)] = c_TCOGS

    # Insert a blank row
    df_cons_Output.loc[len(df_cons_Output.index)] = nan

    # Calculate Gross Margin and Add to DF
    # GM.append('GROSS MARGIN')
    for z in range(mos + 1):
        c_GM.append(c_TR[z + 1] - c_TCOGS[z + 1])
    df_cons_Output.loc[len(df_cons_Output.index)] = c_GM
    
    # Insert a blank row
    df_Output.loc[len(df_Output.index)] = nan

    df_cons_Output.loc[len(df_cons_Output.index)] = 'OpEx'

    for key2, values in cs_oe_dict.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2
    
    # Calculate Total Monthly OpEx (Vertically) and add to DataFrame
    c_TOE = MonthlyTotals(cs_oe_dict, mos, c_TOE)
    df_cons_Output.loc[len(df_cons_Output.index)] = c_TOE

    # Insert a blank row
    df_cons_Output.loc[len(df_cons_Output.index)] = nan

    df_cons_Output.loc[len(df_cons_Output.index)] = 'Non OpEx'
    for key2, values in cs_nonOI_dict.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2

    # Calculate Total Monthly Non OpEx (Vertically) and add to DataFrame
    c_TNOE = MonthlyTotals(cs_nonOI_dict, mos, c_TNOE)
    df_cons_Output.loc[len(df_cons_Output.index)] = c_TNOE

    # Insert a blank row
    df_cons_Output.loc[len(df_cons_Output.index)] = nan

    # Calculate Net Income and Add to DF
    for z in range(mos + 1):
        c_NI.append(c_GM[z + 1] - c_TOE[z + 1] + c_TNOE[z + 1])
    df_cons_Output.loc[len(df_cons_Output.index)] = c_NI

    # Insert a blank row
    df_cons_Output.loc[len(df_cons_Output.index)] = nan

    for z in range(mos + 1):
            EBITDA.append(NI[z + 1] + depreciation[z] + intexpense[z])
        df_Output.loc[len(df_Output.index)] = EBITDA
    #df_cons_Output.to_excel('Cleaned_Consolidated.xlsx', index = False)

    # Output Dataframes to Excel    
    with pd.ExcelWriter('Consolidated Financials.xlsx', engine = 'xlsxwriter') as writer:
        df_cons_Output.to_excel(writer, sheet_name='Consolidated', index = False)
        DAN_df.to_excel(writer, sheet_name='DAN', index = False)
        HQ_df.to_excel(writer, sheet_name='HQ', index = False)
        NYC_df.to_excel(writer, sheet_name='NYC', index = False)
        OAK_df.to_excel(writer, sheet_name='OAK', index = False)
        RMT_df.to_excel(writer, sheet_name='RMT', index = False)
        RWC_df.to_excel(writer, sheet_name='RWC', index = False)
        SF_df.to_excel(writer, sheet_name='SF', index = False)
        SOMA_df.to_excel(writer, sheet_name='SOMA', index = False)
        SV_df.to_excel(writer, sheet_name='SV', index = False)
        VAN_df.to_excel(writer, sheet_name='VAN', index = False)
    #df_consolidated.to_excel('ALL_Sites.xlsx', index = False)
    end = time.time()
    print("The time for this script is:", end-start)
    print('#######')
    print('The Financial Statements have been processed in file "Consolidated Fincancials.xlsx."')
    print('#######')



def FilePrompt():
    root = tk.Tk()
    root.title('Tkinter Open File Dialog')
    root.resizable(False, False)
    root.geometry('300x150')
    root.withdraw()


    filename = fd.askopenfilename()

    return filename

def MonthlyTotals(sec_dict, mos, c_T):

    for c in range(mos + 1):
        t = 0
        for k in sec_dict:
            t = sec_dict[k][c] + t
        c_T.append(t)
    return c_T
    


"""
def PLtoDF(pl_df, months, df_Out):
    print(df_Out)
    for key in pl_df:
        rowlist = [key]
        # Sum up the values of this key, and add to the last value in the dictionary list (Sum horizontially)
        tot = sum(pl_df[key])
        pl_df[key][months] = tot
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for m in pl_df[key]:
            rowlist.append(m)
            df_Out.loc[len(df_Out.index)] = rowlist
    return df_Out
"""

if __name__ == "__main__":
    main()
