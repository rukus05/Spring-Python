import pandas as pd
import datetime
import tkinter as tk
from tkinter import TOP, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.filedialog import asksaveasfile
import PySimpleGUI as sg


def main(): 
    
    
    mos = int(input("How many months are you running this report for? "))
    
    
    
    #########
    # Delete this after troubleshooting
    #########
    #mos = 4
    #moslist = [0] * (mos + 1)
    #########


    revenue_dict={}
    #revenue_dict['Self-pay revenue (Cash, Carrot)'] = [0] * mos
    #revenue_dict['TOTAL REVENUE'] = [0] * mos
    TR = []
    
    COGS_dict={}
    #COGS_dict['TOTAL COGS'] = [0] * mos
    TCOGS = []
    GM = []
    
    oe_dict={}
    #oe_dict['TOTAL OPERATING EXPENSES'] = [0] * mos
    TOE= []
    nonOI_dict = {}
    TNOE = []
    NI = []
    EBITDA = []
    
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
    
    #df_spring['Posting Date'] = df_spring['Posting Date'].dt.date
    df_spring['Posting Date'] = pd.to_datetime(df_spring['Posting Date'])
    df_spring['Posting Date'] = df_spring['Posting Date'].dt.month

    #df_spring.to_excel('test.xlsx', index = False)
    cmos = ['Account']
    for x in range(mos):
        cmos.append('Month ' + str(x + 1))
    cmos.append('Total')
    print(uniqueLoc)
    print(cmos)
    
    #for p in uniquePL:

    for l in uniqueLoc:
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
        #print(revenue_dict)
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

        
        #nonOI_dict['Auto Lease related expenses'] = [0] * (mos + 1)
        #nonOI_dict['Other Income'] = [0] * (mos + 1)
        #nonOI_dict['Interest Income'] = [0] * (mos + 1)
        nonOI_dict['Non-operating income/(expense)'] = [0] * (mos + 1)

        #print(df_Output)
        #print(revenue_dict)
        for index, row in df_spring.iterrows():
            x = str(row['PL Category'])
            m = int(row['Posting Date'])
            #print( str(x) + ' ' + str(m))
            #if ((x == 'Self-pay revenue') or (x == 'Commercial Insurance revenue') or (x == 'Progyny & Stork revenue') or (x == 'Storage revenue' ) or (x == 'Medication') or (x == 'Nest') or (x == 'Other Revenue'))   :
            #    revenue_dict[x][m-1] = revenue_dict[x][m-1] + row['Amount']
            #    print(x)
            #"""
            for key in revenue_dict:
                    
                if (row['PL Category'] == key) and (row['Loc Code Dimension'] == l):
                    r = row['Amount']
                    
                    #print(f'Revenue Dict before is {revenue_dict[key][m-1]}.')
                    revenue_dict[key][m-1] = revenue_dict[key][m-1] + r
                    #print(f'Revenue Dict after is {revenue_dict[key][m-1]}.')
                    ##  The next 2 lines were used for troubleshooting!
                    #if key == 'Self-pay revenue':
                    #    print(f'For {l}, the Key is {key}, month is {m} and amount is {r}.')
                #elif x == 'Medication':
                #    print(x)
                #    revenue_dict[x][m-1] = revenue_dict[row['PL Category']][m-1] + row['Amount']
            #"""
        #print('The revenue dict for all is:')   
        #print(revenue_dict)

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


        df_Output.loc[len(df_Output.index)] = 'Revenue'
        for key in revenue_dict:
            rowlist = [key]
            tot = sum(revenue_dict[key])
            revenue_dict[key][mos] = tot
            for m in revenue_dict[key]:
                rowlist.append(m)
            df_Output.loc[len(df_Output.index)] = rowlist

        

        df_Output.loc[len(df_Output.index)] = 'COGS'
        for key in COGS_dict:
            rowlist = [key]
            tot = sum(COGS_dict[key])
            COGS_dict[key][mos] = tot
            for m in COGS_dict[key]:
                rowlist.append(m)
            df_Output.loc[len(df_Output.index)] = rowlist

        df_Output.loc[len(df_Output.index)] = 'OpEx'
        for key in oe_dict:
            rowlist = [key]
            tot = sum(oe_dict[key])
            oe_dict[key][mos] = tot
            for m in oe_dict[key]:
                rowlist.append(m)
            df_Output.loc[len(df_Output.index)] = rowlist
        df_Output.loc[len(df_Output.index)] = 'Non OpEx'
        for key in nonOI_dict:
            rowlist = [key]
            tot = sum(nonOI_dict[key])
            nonOI_dict[key][mos] = tot
            for m in nonOI_dict[key]:
                rowlist.append(m)
            df_Output.loc[len(df_Output.index)] = rowlist
           
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

        df_Output.to_excel(nm, index = False)              
    #print(revenue_dict)
    # removed nan_df for testing
    df_consolidated = pd.concat([DAN_df, HQ_df, NYC_df, OAK_df, RMT_df, RWC_df, SF_df, SOMA_df, SV_df, VAN_df])
    df_consolidated.to_excel('ALL_Sites.xlsx', index= False)
    print(df_consolidated)

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
