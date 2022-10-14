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

    # Create Dictionary of Lists for COAMapping, as provided by Rosel.

    COA_Dict = {'000-2000 Accounts Payable (A/P)' : 'Accounts Payable', '000-2001 Accounts Payable - FX' : 'Accounts Payable', \
        '000-2003 Accounts Payable (A/P) - USD' : 'Accounts Payable', 'Accounts Payable (A/P) - EUR' : 'Accounts Payable', '000-2585 CC Payable RBC Visa 5771' : 'Accounts Payable', \
        '000-2586 CC Payable RBC Visa USD 9168' : 'Accounts Payable', '000-1086 Due to Sonya' : 'Accounts Payable', '000-2004 RBC Loan #5397-008/005' : 'Accounts Payable', \
        '000-2008 RBC Loan #5397-007/002' : 'Accounts Payable', '000-2009 RBC Loan #5397-004' : 'Accounts Payable', '000-2050 Accrued Expenses' : 'Accounts Payable', \
        '000-2055 Credit Accruals (Patients)' : 'Accounts Payable', '000-2200 PST Payable (Self Assessment)' : 'Accounts Payable', '000-2210 Corporate Tax Payable' : 'Accounts Payable', \
        '000-2300 GST Collected on Sales' : 'Accounts Payable', '000-2410 Deposits - Other' : 'Accounts Payable', 'GST/HST Payable' : 'Accounts Payable', '21000 Account Payable' : 'Accounts Payable', \
        'Accounts Payable (A/P) - EUR' : 'Accounts Payable', 'Accounts Payable (A/P) - USD' : 'Accounts Payable', '22850 RBC Visa xx4230 -CAD' : 'Accounts Payable', \
        '24000 Accrued liabilities' : 'Accounts Payable', '24100 Accrued liabilities:Taxes payable' : 'Accounts Payable', '24600 Accrued liabilities:Other accrued liabilities' : 'Accounts Payable', \
        '24125 PST Payable (Self Assessment)' : 'Accounts Payable', '24150 GST Collected on Sales' : 'Accounts Payable', 'AP' : 'Accounts Payable', 
        '000-2110 Vacation Pay Payable' : 'Accrued PTO', '23400 Payroll liabilities:Accrued PTO' : 'Accrued PTO', \
        '000-2100 Wages Payable' : 'Accrued salaries and wages', '23000 Payroll liabilities' : 'Accrued salaries and wages', \
        '000-3500 Retained Earnings' : 'Accumulated Income/Loss', 'Retained Earnings' : 'Accumulated Income/Loss', \
        '71150 Other operating expenses:Amortization' : 'Amortization', \
        '000-6015 Bank charges' : 'Bank charges', '000-6025 Credit Card Charges' : 'Bank charges', '69110 Bank fees' : 'Bank charges', '69120 Merchant fees' : 'Bank charges', \
        '69130 Payroll processing fee' : 'Bank charges', \
        '000-1031 RBC - Checking - CAD' : 'Cash', '000-1041 RBC - Checking - USD' : 'Cash', '000-1042 RBC USD - FX' : 'Cash', \
        '000-1050 Petty Cash' : 'Cash', '11440 RBC Checking - CAD - 3359' : 'Cash', \
        '200-6605 Physician Salary Expense' : 'Clinical Payroll', '300-6615 Lab Salary Expense' : 'Clinical Payroll', '400-6610 Nurses Salary Expense' : 'Clinical Payroll', '51110 Salaries & wages' : 'Clinical Payroll', \
        '000-2010 Unearned  Revenue' : 'Deferred revenue', '000-2011 Unearned Storage Revenue' : 'Deferred revenue', \
        '000-2220 Future income tax payable' : 'DEFERRED TAX LIABILITY, NET OF CURRENT PORTION', '24175 Accrued liabilities:Taxes payable:Future Income Tax Payable' : 'DEFERRED TAX LIABILITY, NET OF CURRENT PORTION', \
        '000-6030 Depreciation' : 'Depreciation', '71140 Other operating expenses:Depreciation' : 'Depreciation'
        '000-1086 Due from (intercompany):Due from Spring US' : 'Due from Related Parties', '000-1088 Due from (intercompany):Due from Genesis' : 'Due from Related Parties', \
        '000-1082 Due to/from Spring Fertility Management' : 'Due to Related Parties', '000-1083 Due to/from Spring Fertility Vancouver MSO Inc.' : 'Due to Related Parties', \
        '000-1084 Due to/from Spring MSO (contra)' : 'Due to Related Parties', '22200 Management Fees Payable - S.K. MD' : 'Due to Related Parties', '25700 Due to Sonya Kashyap MD Inc.' : 'Due to Related Parties', \
        '25200 Due to (intercompany):Due to SFM' : 'Due to Related Parties', \
        '000-6036 Entertainment - Staff' : 'Employee related expenses', '000-6037 Entertainment - External Party' : 'Employee related expenses', '67110 Professional development' : 'Employee related expenses', \
        '67130 Company events' : 'Employee related expenses', '67160 In-house meals' : 'Employee related expenses', \
        '000-2960 S/H Draws - SK' : 'Equity-Eliminate', '000-3012 Class C Common Shares' : 'Equity-Eliminate', '000-3510 Opening Balance Equity' : 'Equity-Eliminate', \
        '000-3600 Adjustment to Equity' : 'Equity-Eliminate', '000-3610 Adjustment to Equity (Contra)' : 'Equity-Eliminate', '32000 Class C Common Shares' : 'Equity-Eliminate', \
        '100-6020 Computer Software & Support' : 'Facilities', '100-6060 Office - Supplies' : 'Facilities', '100-6061 Office - misc items' : 'Facilities', \
        '100-6065 Repair and maintenance' : 'Facilities', '100-6220 Telephone' : 'Facilities', '200-6840 Membership Dues' : 'Facilities', '400-6047 Janitorial and Laundry - nurse' : 'Facilities', \
        'Freight and delivery - COS' : 'Facilities', '65130 Telephone & internet' : 'Facilities', '65140 Janitorial and waste management' : 'Facilities', '65150 Insurance' : 'Facilities', '65160 Repair & Maintenance' : 'Facilities', \
        '65180 Printing, delivery, & postage' : 'Facilities', '65190 Office supplies' : 'Facilities', '65210 Computer software' : 'Facilities', '65220 Computer hardware' : 'Facilities', '65230 Equipment rentals' : 'Facilities', \
        '65240 Dues & subscriptions' : 'Facilities', 'Insurance' : 'Facilities', \
        '000-6010 Bad Debts' : 'General & administrative', '000-6045 Insurance - Premises' : 'General & administrative', '200-6072 Travel & Conferences - SK' : 'General & administrative', \
        '200-6607 Physicians Health Trust Claim' : 'General & administrative', 'Office expenses' : 'General & administrative', 'Other general and administrative expenses' : 'General & administrative', \
        '000-6920 GST/HST/PST Expense (BC)': 'GST/PST/HST', \
        '17220 Intangible assets:Goodwill' : 'Intangible Assets', '18300 Standard Operating Procedures' : 'Intangible Assets', '18400 License' : 'Intangible Assets', \
        '18500 Non-Hospital Surgical Facility' : 'Intangible Assets', \
        'Interest Expense' : 'Interest', \
        '000-4034 IVF Egg Donor Screening' : 'IVF', '000-4048 IVF Surrogate Screening' : 'IVF', '000-4055 Superovulation Cycle' : 'IVF', '000-4114 Egg Freezing Cycle' : 'IVF', \
        '000-4115 IVF cycle 1' : 'IVF', '000-4120 IVF Cycle 2+' : 'IVF', '000-4132 Sperm Extraction Lab Procedure' : 'IVF', '000-4210 Semen Analysis with morphology' : 'IVF', \
        '000-4211 HBA test' : 'IVF', '000-4216 Sperm DNA Fragmentation Test' : 'IVF', '000-4305 Donor Sperm Insemination Cycle' : 'IVF', '000-4321 Donor Sperm Handling Fee' : 'IVF', \
        '000-4405 IUI Cycle' : 'IVF', '000-4532 Sperm Handling Fee' : 'IVF', '000-4610 Embryo Handling Fee' : 'IVF', '000-4615 Thawing/Replacing Embryos' : 'IVF', '000-4616 Thaw, fertilize eggs, and embryo transfer' : 'IVF', \
        '000-4904 Consults - Dr. Kashyap' : 'IVF', '000-4924 Service Discount Fertile Future' : 'IVF', '000-4935 Non Resident Fee' : 'IVF', '000-4940 Partial Cycle Fee' : 'IVF', '000-4990 Miscellaneous Services' : 'IVF', \
        '000-4999 Write - off' : 'IVF', \
        '000-6870 Service fee - Spring MSO' : 'Management Fee', '41400 Management service revenue' : 'Management Fee', '51400 Management service expense' : 'Management Fee', \
        '000-6205 Advertising and Promotion' : 'Marketing', '63100 Marketing' : 'Marketing', '63120 Events/Swag' : 'Marketing', '63140 Website/multimedia/digital/apps' : 'Marketing', \
        '63150 Advertising, print & promotion' : 'Marketing', \
        '300-6555 PGT-A fee' : 'Medical services', '500-6625 Contract Labour Expenses' : 'Medical services', '51330 Genetic testing' : 'Medical services', '51340 COGS - contract labor' : 'Medical services', \
        '51350 COGS - janitorial and waste management' : 'Medical services', '51370 COGS - repair & maintenance' : 'Medical services', '51380 Laundry' : 'Medical services', \
        '51240 Medication' : 'Medication', \
        '000-4700 Meds - Synarel Nasal Spray' : 'Medications', '000-4701 Meds - Menopur 75IU' : 'Medications', '000-4703 Meds - Endometrin 21 tabs' : 'Medications', '000-4709 Meds - Cetrotide 0.25mg' : 'Medications', \
        '000-4723 Meds - Puregon 300IU' : 'Medications', '000-4729 Meds - Puregon 900IU' : 'Medications', '000-4731 Meds - Pregnyl 10,000 units' : 'Medications', '000-4745 Meds - PPC (hCG) 10,000 iu' : 'Medications', '000-4760 Meds - Orgalutran 0.25 mg' : 'Medications', \
        '000-4761 Estradot Patch 100mcg/24H' : 'Medications', '000-4770 Meds - Prometrium 100 mg Caps' : 'Medications', '000-4783 Meds - Marvelon Tabs 21' : 'Medications', '000-4784 Meds - Decapeptyl 1ML' : 'Medications', '000-4801 Pharmacy - Male Vitamins' : 'Medications', \
        '000-4804 Pharmacy - CoQ10' : 'Medications', \
        '000-4994 MSP Billing' : 'MSP', \
        '000-6090 Miscellaneous' : 'Other', '70000 Other operating expenses' : 'Other', '71170 Other operating expenses:Other misc expense' : 'Other', 'Exchange Gain or Loss' : 'Other', \
        '000-1063 CEWS receivable' : 'Other Current Assets', '000-1078 Accounts receivable - income taxes' : 'Other Current Assets', '000-1068 Inventory - Medications' : 'Other Current Assets', '000-1069 Inventory - Pharm Misc Items' : 'Other Current Assets', \
        '000-1070 Prepaid Expense' : 'Other Current Assets', '000-1071 Prepaids - Other' : 'Other Current Assets', '000-1085 Goodwill' : 'Other Current Assets', '000-1066 GST Paid' : 'Other Current Assets', '000-1068 Inventory:Medication' : 'Other Current Assets', \
        '000-1070 Prepaids' : 'Other Current Assets', '000-1071 Prepaids:Prepaid insurance' : 'Other Current Assets', '000-1075 Prepaids:Other prepaid expenses' : 'Other Current Assets', \
        '30000 Contingent Liabilites' : 'Other Long Term Liabilities', \
        '000-1080 Accounts Receivable - GFC Ent.' : 'Other Receivables', '000-1064 SRED receivable' : 'Other Receivables', \
        '000-4991 Admin Photocopying Fees' : 'Other Revenue', '000-4992 Space rent' : 'Other Revenue', \
        '000-1060 Accounts Receivable (A/R)' : 'Patient accounts receivable, net of alowance for doubtful accounts', '000-1061 Accounts Receivable (A/R) - USD' : 'Patient accounts receivable, net of alowance for doubtful accounts', \
        '000-1062 Allowance for Doubtful Accounts', '000-1065 A/R' : 'Patient accounts receivable, net of alowance for doubtful accounts', \
        '100-6620 Office Salary Expense' : 'Payroll', '100-6630 IT salary expense' : 'Payroll', '51114 Cost of Goods Sold' : 'Payroll', '61110 Salaries & wages' : 'Payroll',  \
        '000-6825 Accounting' : 'Professional Fees', '000-6830 Legal fees' : 'Professional Fees', '64110 Accounting' : 'Professional Fees', '64120 Consulting' : 'Professional Fees', '64130 Contract labor' : 'Professional Fees', '64160 Legal' : 'Professional Fees', \
        '000-1800 Office Furniture' : 'Property and Equipment, net', '000-1810 Acc. Depre - Office Furniture' : 'Property and Equipment, net', '000-1820 Office Equipment' : 'Property and Equipment, net', \
        '000-1830 Acc. Depre - Office Equipment' : 'Property and Equipment, net', '000-1840 Computer Equipment' : 'Property and Equipment, net', '000-1850 Acc. Depre - Computer Equipment' : 'Property and Equipment, net', '000-1860 Medical Equipment' : 'Property and Equipment, net', \
        '000-1870 Acc. Depre - Medical Equipment' : 'Property and Equipment, net', '000-1880 Software' : 'Property and Equipment, net', '000-1890 Acc. Depre - Software' : 'Property and Equipment, net', '000-1955 Leasehold Imprv. Broadway' : 'Property and Equipment, net', \
        '000-1956 Acc. Depre - Leasehold Imprv. Broadway' : 'Property and Equipment, net', '16210 Fixed assets:Medical & lab  equipment - net:Medical & lab equipment - gross' : 'Property and Equipment, net', \
        '16220 Fixed assets:Medical & lab  equipment - net:Accumulated depreciation - medical & lab equipment' : 'Property and Equipment, net', \
        '16310 Fixed assets:IT equipment - net:IT equipment - gross' : 'Property and Equipment, net', '16320 Fixed assets:IT equipment - net:Accumulated depreciation - IT  equipment' : 'Property and Equipment, net', \
        '16410 Fixed assets:Leasehold improvements - net:Leasehold improvements - gross' : 'Property and Equipment, net', \
        '16420 Fixed assets:Leasehold improvements - net:Accumulated depreciation - leasehold improvements' : 'Property and Equipment, net', '17100 Fixed assets:Software - net:Software - gross' : 'Property and Equipment, net', \
        '17110 Fixed assets:Software - net:Accumulated amortization - software' : 'Property and Equipment, net', \
        '65110 Base rent' : 'Rent', \
        '11460 GIC' : 'Restricted Cash', \
        '000-3700 Distribution' : 'Shareholder Distribution', '3700 Draw' : 'Shareholder Distribution', \
        '000-4520 Donor Sperm Annual Storage' : 'Storage', '000-4525 Elective Sperm Freezing' : 'Storage', '000-4526 IVF Sperm Freeze Backup' : 'Storage', '000-4530 Frozen Sperm Annual Storage' : 'Storage', \
        '000-4620 Frozen Embryo Annual Storage Fee' : 'Storage', \
        '300-6210 Couriers & Freight' : 'Supplies', '300-6212 Brokerage Fees' : 'Supplies', '300-6410 Lab Supplies - Consumables' : 'Supplies', '300-6411 Lab supplies - Gases' : 'Supplies', \
        '400-6415 Nursing Supplies' : 'Supplies', '400-6420 Pharmaceutical - Nursing' : 'Supplies', '400-6425 Pharmaceuticals - Meds' : 'Supplies', '51200 COGS - supplies' : 'Supplies', '51210 Surgical supplies' : 'Supplies', \
        '51220 Clinical supplies' : 'Supplies', '51230 IVF Lab supplies' : 'Supplies', \
        'Taxes and Licenses' : 'Taxes & regulatory', '68100 Taxes & Regulatory fees' : 'Taxes & regulatory', '68130 Licensing & permitting' : 'Taxes & regulatory', \
        '300-6075 Travel, Conferences, Education' : 'Travel', '400-6075 Travel, Conferences, Education - Nursing' : 'Travel', '66110 Airfare' : 'Travel', '66130 Meals' : 'Travel', \
        '66140 Ground transportation' : 'Travel', '66160 Other travel expenses' : 'Travel', 'Meals and entertainment' : 'Travel'}
    

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
    cs_EBITA = {}
    
    
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
        sttax  = [0] * (mos + 1)
        fedtax = [0] * (mos + 1)
        EBITA = []
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
        EBITA.append('Monthly EBITA')
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
            
            if (row.Loc_Code_Dimension == l) and (row.GL_Account == '68110State and local taxes'):
                sttax[m-1] = sttax[m-1] + row.Amount
            if (row.Loc_Code_Dimension == l) and (row.GL_Account == '68120Federal taxes'):
                fedtax[m-1] = fedtax[m-1] + row.Amount
 
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
            EBITA.append(NI[z + 1] + depreciation[z] + intexpense[z] + sttax[z] + fedtax[z])
        df_Output.loc[len(df_Output.index)] = EBITA
        

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
    cs_EBITA['Monthly EBITA'] = [0] * (mos + 1)

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
        elif row[0] in cs_EBITA:
            b = row[0]
            for a in range(mos + 1):
                cs_EBITA[b][a] = (cs_EBITA[b][a]) + (row[a + 1])

    #print(cs_revenue_dict)
    
    # Create new DF to hold cleaned dataframe
    df_cons_Output = pd.DataFrame(columns= cmos)
    
    c_TR = []
    c_TCOGS = []
    c_GM = []
    c_TOE= []
    c_TNOE = []
    c_NI = []
    #c_EBITDA = []

    c_TR.append('Monthly Total Revenue')
    c_TCOGS.append('Monthly Total COGS')
    c_GM.append('Monthly GROSS MARGIN')
    c_TOE.append('Monthly Total OpEx')
    c_TNOE.append('Monthly Total Non OpEx')
    c_NI.append('Monthly Net Income')
    #cs_EBITA.append('Monthly EBITA')
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

    for key2, values in cs_EBITA.items():
        rowlist2 = [key2]
        #  Append all values of the dictionary key to the List and insert into the dataframe
        for v in values:
            rowlist2.append(v)
            #print(rowlist2)
        df_cons_Output.loc[len(df_cons_Output.index)] = rowlist2

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
    runningtime = time.time() - start
    print("The time for this script is:", runningtime)
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
