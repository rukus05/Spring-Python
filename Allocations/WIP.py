from GetTables import create_empalloc_dict
from GetTables import create_deptalloc_dict
from GetTables import deptcode_to_subdept as dts
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


def main():

    start = time.time()


    ###  Section Allocations ###

    # Prompt user for the Allocations data
    print("Select the current allocations File:")
    allocf = FilePrompt()
    df_ea = pd.read_excel(allocf)
    df_ea = df_ea.reset_index()

    # Pass dataframe to create_empalloc_dict function to create employee allocations dictionary
    emp_alloc_dict = create_empalloc_dict(df_ea)
    # Pass dataframe to create_deptalloc_dict function to create department allocations dictionary
    dept_alloc_dict = create_deptalloc_dict(df_ea)
    



    # Open input file
    print("Select the current Input File:")
    inputf = FilePrompt()
    df = pd.read_excel(inputf)
    df = df.reset_index()

    # Fill in blank cells with 0
    df['GROSS PAY less PTO USED, Bonus, OT'] = df['GROSS PAY less PTO USED, Bonus, OT'].fillna(0)
    df['OT'] = df['OT'].fillna(0)
    df['VOLUNTARY DEDUCTION : ELC-ELECTRONICS RMB'] = df['VOLUNTARY DEDUCTION : ELC-ELECTRONICS RMB'].fillna(0)
    df['TOTAL EMPLOYER TAX'] = df['TOTAL EMPLOYER TAX'].fillna(0)
    df['MEMO : KM-401K SH MATCH'] = df['MEMO : KM-401K SH MATCH'].fillna(0)


    # Create new Dataframe for the Allocations Output.
    df_allocations = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'AcctType', 'AcctNo', 'AcctName', 'Description', \
                                           'DebitAmt', 'CreditAmt', 'Loc', 'Dept','Provider', 'Service Line', 'Comments'])
    
    # Create a list of all Locations
    all_locations = ['HQ', 'Nest', 'SF', 'OAK', 'SV', 'NYC', 'PDX']
    # Create a list of all values to allocate
    all_values = ['GROSS PAY less PTO USED, Bonus, OT', 'OT', 'VOLUNTARY DEDUCTION : ELC-ELECTRONICS RMB', 'TOTAL EMPLOYER TAX', \
                     'MEMO : KM-401K SH MATCH']

    for index, row in df.iterrows():
        pid = row['POSITION ID']
        dept = row['Department']
        # Check if dept is Receptionist HQ. If so, check if there is a seperate allocation for the employee. (2nd if)
        # If so, use the % as defined in the dictionary when create_empalloc_dict is called.
        # If not, use the % as defined in the dictionary when create_deptalloc_dict is called.
        # This process repeats for other depts.  If we don't match any of these, use the emp allocation; last else statement below.
        if dept == 'Receptionist HQ':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['HQ']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['HQ']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        
        elif dept == 'Medical Records':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['HQ']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['HQ']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        elif dept == 'Call Center':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['HQ']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['HQ']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        elif dept == 'Financial Counselor':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['HQ']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['HQ']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        elif dept == 'Clinical Operations':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['HQ']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['HQ']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        elif pid in emp_alloc_dict:
            hq_percent = emp_alloc_dict[pid]['HQ']
            print(hq_percent)
            nest_percent = emp_alloc_dict[pid]['Nest']
            sf_percent = emp_alloc_dict[pid]['SF']
            oak_percent = emp_alloc_dict[pid]['OAK']
            sv_percent = emp_alloc_dict[pid]['SV']
            nyc_percent = emp_alloc_dict[pid]['NYC']
            pdx_percent = emp_alloc_dict[pid]['PDX']

        # Iterate through all locations.  This calculates the allocations, and creates a line in the dataframe for each location.
        # 
        #  df_allocations = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'AcctType', 'AcctNo', 'AcctName', 'Description', \
        #                                        'DebitAmt', 'CreditAmt', 'Loc', 'Dept','Provider', 'Service Line', 'Comments'])

        for v in all_values:
            if row[v] != 0.0:
                df_allocations.loc[len(df_allocations.index)] = [row['COMPANY CODE'], row['PERIOD ENDING DATE'], row['PAY DATE'], 'G/L Account', \
                                                                'AcctNo', ' ', 'Description', ' ', row[v], row['Office Reporting Location'], 'Dept', \
                                                                'NULL', 'NULL', 'Comments']
                for l in all_locations:
                    if l == 'HQ':
                        pct = hq_percent
                    elif l == 'Nest':
                        pct = nest_percent
                    elif l == 'SF':
                        pct = sf_percent
                    elif l == 'OAK':
                        pct = oak_percent
                    elif l == 'SV':
                        pct = sv_percent
                    elif l == 'NYC':
                        pct = nyc_percent
                    elif l == 'PDX':
                        pct = pdx_percent
                    
                    if pct != 0.0:
                        allocated_value = row[v]*pct
                        df_allocations.loc[len(df_allocations.index)] = [row['COMPANY CODE'], row['PERIOD ENDING DATE'], row['PAY DATE'], 'G/L Account', \
                                                                'AcctNo', ' ', 'Description', allocated_value , ' ', l, 'Dept', \
                                                                'NULL', 'NULL', 'Comments']

            
     # Start the "Save As" dialog box.
    app = tk.Tk()
    app.title("Save File As")
    status_label = tk.Label(app, text="", fg="green")
    status_label.pack()
    save_button = tk.Button(app, text="Save as", command=save_dataframe(df_allocations, status_label))
    save_button.pack(padx=20, pady=10)

    # Calculate the execution time.
    runningtime = time.time() - start
    print("The execution time is:", runningtime)


def FilePrompt():
    root = tk.Tk()
    root.title('Tkinter Open File Dialog')
    root.resizable(False, False)
    root.geometry('300x150')
    root.withdraw()


    filename = fd.askopenfilename()

    return filename
    
def save_dataframe(df, sl):
    file_path = fd.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    
    if file_path:
        try:
            # Assuming df is your DataFrame
            df.to_excel(file_path, index=False)
            sl.config(text=f"Saved as {file_path}")
        except Exception as e:
            sl.config(text=f"Error: {str(e)}")


if __name__ == "__main__":
    main()
