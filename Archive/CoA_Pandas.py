

from msilib.schema import File
import pandas as pd
import openpyxl
import datetime
import tkinter as tk
from tkinter import TOP, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.filedialog import asksaveasfile
import PySimpleGUI as sg


def main(): 
    
    # For this program to work properly, the CoA_Pandas.py file needs to be in the same directory as the RawData.xlsx file.
    # Read in Data from the "RawData.xlsx" file.
    f = FilePrompt()
    df_spring = pd.read_excel(f)
    #df_spring = pd.read_excel('Allocated Data Output.xlsx')
    
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

    # Create new Dataframe for the Output.
    df_Output = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', 'DebitAmt', 'CreditAmt', 'Loc', 'Dept', 'Provider', 'Service Line', 'Comments'])
    # print(df_Output)

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
                    for index, row in df_spring.iterrows():
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
                        if TotClientCharges_Sum != 0:  
                            df_Output.loc[len(df_Output.index)] = ["", ped.date(), ivd.date(), "", str(k), CoA[cnt][CoA_Index], "", str(i) + ' ' + str(j), x, "", l, deptCode, "", "", ""]
                            cnt = cnt + 1       # Increment CoA counter
                            
                    # Add a credit entry to match the above debits.
                    if TotClientCharges_Sum != 0:     
                        df_Output.loc[len(df_Output.index)] = [ "", ped.date(), ivd.date(), "", str(k), 23300, "", str(i) + ' ' + str(j), "", TotClientCharges_Sum , l, deptCode, "", "", ""]
                    
                   
                    # Test Code
                    #print(row['Invoice Number'], row['Department Long Descr'], row['SUB_DEPARTMENT'], row['LOCATION'], GrossWages_Sum, Roth401kCombo_Sum)
    

    
        op = input("Please type name of file for Output:")
        output = str(op + '.xlsx')
        df_Output.to_excel(output, index = False)

        
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



