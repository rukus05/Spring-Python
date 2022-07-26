

import pandas as pd
import openpyxl
import datetime
import tkinter as tk
from tkinter import TOP, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.filedialog import asksaveasfile


def main(): 
    
    # For this program to work properly, the CoA_Pandas.py file needs to be in the same directory as the RawData.xlsx file.
    # Read in Data from the "RawData.xlsx" file.
    f = FilePrompt()
    df_spring = pd.read_excel(f)
    

    
    df_spring = df_spring.reset_index()
    
    # Get the unique Invoice Numbers, Locations, Sub Departments, and Department Long Descriptions.  This is needed to loop through and group them properly.
    # uniqueInvoices = df_spring['Invoice Number'].unique()
    unique_Locations = df_spring['LOCATION'].unique()
    unique_SubDept = df_spring['SUB_DEPARTMENT'].unique()
    unique_DptCode = df_spring['DEPT CODE'].unique()
    unique_Entity = df_spring['ENTITY'].unique()
    #print(unique_DLD)
    #print(df_spring.info())
    #  It's important to fill in blank cells for the below columns with Zeros.  A blank cell breaks the calculations
    #  The .fillna() method fills blank cells in these columns with 0's.
    df_spring['ACCRUAL'] = df_spring['ACCRUAL'].fillna(0)
    df_spring['ENTITY'] = df_spring['ENTITY'].fillna(0)
    
    
    # Create new Dataframe for the Output.
    df_Output = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', 'DebitAmt', 'CreditAmt', 'Loc', 'Dept', 'Provider', 'Service Line', 'Comments'])
    # print(df_Output)

    # By using 4 nested FOR loops, we can group the rows by Invoices, Dept Long Desc, Sup Dept, and Location
    for i in unique_SubDept:
        for j in unique_Locations:
            for k in unique_DptCode:
                # Initialize Sum Variables
                Accrual_Sum = 0
                                
                # Loop through rows of the Raw Data file.
                for index, row in df_spring.iterrows():
                    # These if statements force a match for specific Invoices, Dept Long Desc, Sup Dept, and Locations
                    if (row['SUB_DEPARTMENT'] == i) and (row['LOCATION'] == j) and (row['DEPT CODE'] == k):
                        # Sum up the pertinent columns.
                        Accrual_Sum = Accrual_Sum + row['ACCRUAL']
                        glCode = row['GL CODE']
                        deptCode = row['DEPT CODE']
                        #subDept = row['SUB_DEPARTMENT']
                       
                       

                # Create an array holding the Sums for a particular Invoice, DLD, Location, and Sub_Dept        
                summary = [i, j, k , Accrual_Sum]          
                                
                
                # Initialize counter for CoA
                #cnt = 4
                # Loop over the Sum variables in the Summary Array.  
                # Add row to output file if GrossWages doesn't equal 0.  
                # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                
                if Accrual_Sum != 0:  
                    df_Output.loc[len(df_Output.index)] = ["", "", "", "", i, glCode, "", "", Accrual_Sum, "", j, deptCode, "", "", ""]
                #   cnt = cnt + 1       # Increment CoA counter
                    
                # Add a credit entry to match the above debits.
                #if GrossWages_Sum != 0:     
                #    df_Output.loc[len(df_Output.index)] = [ "", ped.date(), ivd.date(), "", str(k), 23300, "", str(i) + ' ' + str(j), "", TotClientCharges_Sum , l, deptCode, "", "", ""]
                
                
                # Test Code
                #print(row['Invoice Number'], row['Department Long Descr'], row['SUB_DEPARTMENT'], row['LOCATION'], GrossWages_Sum, Roth401kCombo_Sum)


    #  Export the Pandas DataFrame to an Excel file called "Output_test.xlsx"
    inp = input("Please type name of file for Output:")
    des = str(inp + '.xlsx')
    df_Output.to_excel(des, index = False)
    

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



