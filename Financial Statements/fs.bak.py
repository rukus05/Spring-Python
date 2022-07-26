import pandas as pd
import datetime
import tkinter as tk
from tkinter import TOP, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.filedialog import asksaveasfile
import PySimpleGUI as sg


def main(): 
    

    mo = input("How many months are you running this report for? ")
    mos = int(mo) + 1
    moslist = list(range(mos))

    # For this program to work properly, the CoA_Pandas.py file needs to be in the same directory as the RawData.xlsx file.
    # Read in Data from the "RawData.xlsx" file.
    

    revenue_dict={}
    revenue_dict['Self-pay revenue (Cash, Carrot)'] = list(range(mos))
    revenue_dict['Commercial Insurance revenue'] = list(range(mos))
    revenue_dict['Progyny & Stork revenue'] = list(range(mos))
    revenue_dict['Storage revenue'] = list(range(mos))
    revenue_dict['Medication'] = list(range(mos))
    revenue_dict['Nest'] = list(range(mos))
    revenue_dict['Other Revenue'] = list(range(mos))
    #revenue_dict['TOTAL REVENUE'] = list(range(mos))
    TR = []
    print(revenue_dict)
    COGS_dict={}
    COGS_dict['MD Payroll'] = list(range(mos))
    COGS_dict['Clinical Payroll'] = list(range(mos))
    COGS_dict['Lab Payroll'] = list(range(mos))
    COGS_dict['ASC Payroll'] = list(range(mos))
    COGS_dict['Supplies'] = list(range(mos))
    COGS_dict['Medication'] = list(range(mos))
    COGS_dict['Medical Services'] = list(range(mos))
    #COGS_dict['TOTAL COGS'] = list(range(mos))
    TCOGS = []
    GM = []
    print(COGS_dict)
    oe_dict={}
    oe_dict['Payroll'] = list(range(mos))
    oe_dict['Marketing'] = list(range(mos))
    oe_dict['Professional Fees'] = list(range(mos))
    oe_dict['Rent'] = list(range(mos))
    oe_dict['Facilities'] = list(range(mos))
    oe_dict['Travel'] = list(range(mos))
    oe_dict['Facilities'] = list(range(mos))
    oe_dict['Employee Related Expenses'] = list(range(mos))
    oe_dict['Travel & Reguatory'] = list(range(mos))
    oe_dict['Bank Charges'] = list(range(mos))
    oe_dict['Other'] = list(range(mos))
    #oe_dict['TOTAL OPERATING EXPENSES'] = list(range(mos))
    TOE= []
    nonOI = {}
    nonOI['Auto Lease related expenses'] = list(range(mos))
    nonOI['Other Income'] = list(range(mos))
    nonOI['Interest Income'] = list(range(mos))
    TNOE = []
    NI = []
    EBITDA = []
    
    
    f = FilePrompt()
    df_spring = pd.read_excel(f)
    
    df_spring = df_spring.reset_index()
    df_spring['Amount'] = df_spring['Amount'].fillna(0)
    
    uniquePL = df_spring['PL Category'].unique()
    #df_spring['Posting Date'] = df_spring['Posting Date'].dt.date
    df_spring['Posting Date'] = pd.to_datetime(df_spring['Posting Date'])
    df_spring['Posting Date'] = df_spring['Posting Date'].dt.month

    df_spring.to_excel('test.xlsx', index = False)
    
    for p in uniquePL:
        for index, row in df_spring.iterrows():
            for m in range(mos):
                if row['Month No'] == m:
                    print (row['Posting Date'])

    

    
    """
    df['date'] = pd.to_datetime(df['date'])
    df[df['date'].dt.month > 2]
    """

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
