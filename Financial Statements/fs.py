from math import nan
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
    df_spring['Amount'] = df_spring['Amount'].fillna(0)
    
    uniquePL = df_spring['PL Category'].unique()
    uniqueLoc = df_spring['Loc Code Dimension'].unique()
    
    
    df_spring['Posting Date'] = pd.to_datetime(df_spring['Posting Date'])
    df_spring['Posting Date'] = df_spring['Posting Date'].dt.month

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

        # Iterate through df_spring DF
        for index, row in df_spring.iterrows():
            x = str(row['PL Category'])
            m = int(row['Posting Date'])
            
            # Based on location, add the PL values.
            # A for loop for each section of the finanancial statement (dict)
            for key in revenue_dict:
                if (row['PL Category'] == key) and (row['Loc Code Dimension'] == l):
                    r = row['Amount']
                    revenue_dict[key][m-1] = revenue_dict[key][m-1] + r
            for key in COGS_dict:
                if (row['PL Category'] == key) and (row['Loc Code Dimension'] == l):
                    r = row['Amount']
                    COGS_dict[key][m-1] = COGS_dict[key][m-1] + r
            for key in oe_dict:
                if (row['PL Category'] == key) and (row['Loc Code Dimension'] == l):
                    r = row['Amount']
                    oe_dict[key][m-1] = oe_dict[key][m-1] + r

            for key in nonOI_dict:
                if (row['PL Category'] == key) and (row['Loc Code Dimension'] == l):
                    r = row['Amount']
                    nonOI_dict[key][m-1] = nonOI_dict[key][m-1] + r
        
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
        for c in range(mos + 1):
            t = 0
            for k in revenue_dict:
                t = revenue_dict[k][c] + t
            TR.append(t)
        #print(TR)
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
        for c in range(mos + 1):
            t = 0
            for k in COGS_dict:
                t = COGS_dict[k][c] + t
            TCOGS.append(t)
        #print(TR)
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
        for c in range(mos + 1):
            t = 0
            for k in oe_dict:
                t = oe_dict[k][c] + t
            TOE.append(t)
        #print(TR)
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
        for c in range(mos + 1):
            t = 0
            for k in nonOI_dict:
                t = nonOI_dict[k][c] + t
            TNOE.append(t)
        #print(TR)
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
        #
        #for z in range(mos + 1):
        #    EBITDA.append(NI[z + 1] - TOE[z + 1] + TNOE[z + 1])
        #df_Output.loc[len(df_Output.index)] = NI
        

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
    print(df_cons_Output)
    print(cs_revenue_dict)

    # Place Consolidate Revenue Account info into DF
    for key2, values in cs_revenue_dict.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2
    
    for key2, values in cs_COGS_dict.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2

    for key2, values in cs_oe_dict.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2
    
    for key2, values in cs_nonOI_dict.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2
        
    df_cons_Output.to_excel('Cleaned_Consolidated.xlsx', index = False)
        

    


    with pd.ExcelWriter('Consolidated Financials.xlsx', engine = 'xlsxwriter') as writer:
        df_consolidated.to_excel(writer, sheet_name='Consolidated', index = False)
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
