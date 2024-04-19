from GetTables import create_empalloc_dict
from GetTables import create_deptalloc_dict
from GetTables import deptcode_to_subdept
from GetTables import entity_tagging
from GetTables import chart_of_accounts
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
    #print(emp_alloc_dict)
    # Pass dataframe to create_deptalloc_dict function to create department allocations dictionary
    dept_alloc_dict = create_deptalloc_dict(df_ea)
    print(dept_alloc_dict)
    
    # Prompt user for file containg Dept Code to Sub Dept mappings
    print("Select the current Dept Code to Sub Dept Mappings File:")
    mappingsf = FilePrompt()
    deptcodetosubmappings_df = pd.read_excel(mappingsf)
    deptcodetosubmappings_df = deptcodetosubmappings_df.reset_index()
    # Pass the dept code submappings dataframe to the deptcode_to_subdept function to create the dept code 
    # to sub dept dictionary
    deptcodetosub_dict = deptcode_to_subdept(deptcodetosubmappings_df)
    #print(deptcodetosub_dict)
    # Prompt user for the entity tagging file.
    print("Select the Entity Tagging File:")
    entityf = FilePrompt()
    entitytagging_df = pd.read_excel(entityf, dtype=str)
    entitytagging_df = entitytagging_df.reset_index()
    # Pass the entity tagging dataframe to the deptcode_to_subdept function to create the dept code
    entitytagging_dict = entity_tagging(entitytagging_df)
    #print(entitytagging_dict)

    # Prompt user for Chart of Accounts File
    print("Select the Chart of Accounts File:")
    coaf = FilePrompt()
    coa_df = pd.read_excel(coaf)
    coa_df = coa_df.reset_index()
    coa_dict = chart_of_accounts(coa_df)

    #print(coa_dict)

    # Prompt user for Input file
    print("Select the Input File which your running Payroll Allocations for:")
    inputf = FilePrompt()
    df = pd.read_excel(inputf,dtype={'FILE NUMBER': str})
    df = df.reset_index()

    # Fill in blank cells with 0
    df = df.fillna(0)
    #df['GROSS PAY less PTO USED, Bonus, OT'] = df['GROSS PAY less PTO USED, Bonus, OT'].fillna(0)
    #df['OT'] = df['OT'].fillna(0)
    #df['VOLUNTARY DEDUCTION : ELC-ELECTRONICS RMB'] = df['VOLUNTARY DEDUCTION : ELC-ELECTRONICS RMB'].fillna(0)
    #df['TOTAL EMPLOYER TAX'] = df['TOTAL EMPLOYER TAX'].fillna(0)
    #df['MEMO : KM-401K SH MATCH'] = df['MEMO : KM-401K SH MATCH'].fillna(0)
    #df['Medical Waiver'] = df['Medical Waiver'].fillna(0)
    #df['MEDICAL'] = df['MEDICAL'].fillna(0)
    #df['DENTAL'] = df['DENTAL'].fillna(0)
    #df['VISION'] = df['VISION'].fillna(0)
    #df['LIFE'] = df['LIFE'].fillna(0)
    #df['MEDICAL'] = df['MEDICAL'].fillna(0)

    # Create new Dataframe for the Employee Allocations Output.
    df_emp_allocations = pd.DataFrame(columns=['Entity Template', 'Entity', 'PostDate', 'DocDate', 'DocNo','AcctType', 'AcctNo', 'AcctName', 'Description', \
                                           'DebitAmt', 'CreditAmt', 'Loc', 'Dept','Provider', 'Service Line', 'Comments'])
    # Create new Dataframe for the Dept Allocations Output.
    df_dept_allocations = pd.DataFrame(columns=['Entity Template', 'Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', \
                                           'DebitAmt', 'CreditAmt', 'Loc', 'Dept','Provider', 'Service Line', 'Comments'])
    
    # Capture home payroll company of the file.
    if str(df.loc[1, 'COMPANY CODE'])  == '362':
        hc = '362'
        ent_template = '008'
    elif str(df.loc[1, 'COMPANY CODE'])  == '22J':
        hc = '22J'
        ent_template = '007'
    elif str(df.loc[1, 'COMPANY CODE']) == 'ML7':
        hc = 'ML7'
        ent_template = '002'
    
    #print(hc)
    # Create a list of all Locations
    all_locations = ['SFM MSO', 'Nest', 'SF', 'OAK', 'SV', 'NYC', 'PDX']
    # Create list for ALlocated depts
    all_alloc_depts = list(dept_alloc_dict.keys())
    print(all_alloc_depts)
    #all_alloc_depts = ['Receptionist HQ', 'Medical Records', 'Call Center', 'Financial Counselor', 'Clinical Operations', 'Revenue Cycle']
    # Create a list of all values to allocate
    coa_headers = coa_df.columns
    all_values = coa_headers.tolist()

    # Create a list of the allocations file headers.  
    alloc_headers = df.columns
    alloc_headers_values = alloc_headers.tolist()

    # Create a set to capture Allocation headers that are not in the input file
    missing_headers = []

    #all_values = ['GROSS PAY less PTO USED, Bonus, OT', 'OT', 'VOLUNTARY DEDUCTION : ELC-ELECTRONICS RMB', 'TOTAL EMPLOYER TAX', \
    #                 'MEMO : KM-401K SH MATCH']c
    #print(all_values)


    for index, row in df.iterrows():
        pid = str(row['POSITION ID'])
        #print(pid)
        dept = row['Department']
        cc = row['COMPANY CODE']
        # Intialize employee percentages variables
        emp_hq_percent = 0
        emp_nest_percent = 0
        emp_sf_percent = 0
        emp_oak_percent = 0
        emp_sv_percent = 0
        emp_nyc_percent = 0
        emp_pdx_percent = 0
        # Intialize department percentages variables
        dept_hq_percent = 0
        dept_nest_percent = 0
        dept_sf_percent = 0
        dept_oak_percent = 0
        dept_sv_percent = 0
        dept_nyc_percent = 0
        dept_pdx_percent = 0
        # Check if dept is Receptionist HQ. If so, check if there is a seperate allocation for the employee. (2nd if)
        # If so, use the % as defined in the dictionary when create_empalloc_dict is called.
        # If not, use the % as defined in the dictionary when create_deptalloc_dict is called.
        # This process repeats for other depts.  If we don't match any of these, use the emp allocation; last else statement below.
        if re.search('Receptionist HQ*', str(dept), re.IGNORECASE):
            #dept == 'Receptionist HQ':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['SFM MSO']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['SFM MSO']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        
        elif re.search('Medical Records*', str(dept), re.IGNORECASE):
             #dept == 'Medical Records':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['SFM MSO']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['SFM MSO']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        elif re.search('Call Center*', str(dept), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Call Center '):
            # Accommodate Allocations file that has Call Center as lower case.
            dept = 'Call Center'
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['SFM MSO']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['SFM MSO']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']

                print('Call Center Hit, ' + str(sf_percent))
        elif re.search('Financial Counselor*', str(dept), re.IGNORECASE):
            #dept == 'Financial Counselor':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['SFM MSO']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['SFM MSO']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        elif re.search('Clincal Operations*', str(dept), re.IGNORECASE):
            # dept == 'Clinical Operations':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['SFM MSO']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['SFM MSO']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        elif re.search('Revenue Cycle*', str(dept), re.IGNORECASE):
            # dept == 'Revenue Cycle':
            if pid in emp_alloc_dict:
                hq_percent = emp_alloc_dict[pid]['SFM MSO']
                nest_percent = emp_alloc_dict[pid]['Nest']
                sf_percent = emp_alloc_dict[pid]['SF']
                oak_percent = emp_alloc_dict[pid]['OAK']
                sv_percent = emp_alloc_dict[pid]['SV']
                nyc_percent = emp_alloc_dict[pid]['NYC']
                pdx_percent = emp_alloc_dict[pid]['PDX']
            else:
                hq_percent = dept_alloc_dict[dept]['SFM MSO']
                nest_percent = dept_alloc_dict[dept]['Nest']
                sf_percent = dept_alloc_dict[dept]['SF']
                oak_percent = dept_alloc_dict[dept]['OAK']
                sv_percent = dept_alloc_dict[dept]['SV']
                nyc_percent = dept_alloc_dict[dept]['NYC']
                pdx_percent = dept_alloc_dict[dept]['PDX']
        elif pid in emp_alloc_dict:
            hq_percent = emp_alloc_dict[pid]['SFM MSO']
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

        for v in all_values[2:]:
            # This commented out item bypassed checking for all the items in the COA file.
            if v in alloc_headers_values:
                if row[v] != 0.0:
                    if pid in emp_alloc_dict:  # if employee has allocation, use that allocation.  if not, use department allocatioin 
                        df_emp_allocations.loc[len(df_emp_allocations.index)] = [ent_template, entitytagging_dict[hc][str(row['Office Reporting Location'])], row['PERIOD ENDING DATE'], row['PAY DATE'], ' ', 'G/L Account', \
                                                                        str(coa_dict[row['Sub Department']][v]), ' ', str(row['COMPANY CODE']) + '-' + str(row['PERIOD ENDING DATE']) + '-' + dept + '-' + v + '-' + row['Sub Department'] + '-' + row['Office Reporting Location'] + '-' + pid, \
                                                                        ' ', row[v], row['Office Reporting Location'], '0' + str(deptcodetosub_dict[row['ADP Department Code']]), \
                                                                        'NULL', 'NULL', str(row['COMPANY CODE']) + '- Allocations - PPE ' + row['PERIOD ENDING DATE']]
                        for l in all_locations:
                            if l == 'SFM MSO':
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
                                df_emp_allocations.loc[len(df_emp_allocations.index)] = [ent_template, entitytagging_dict[str(row['COMPANY CODE'])][l], row['PERIOD ENDING DATE'], row['PAY DATE'], ' ', 'G/L Account', \
                                                                        coa_dict[row['Sub Department']][v], ' ', str(row['COMPANY CODE']) + '-' + str(row['PERIOD ENDING DATE']) + '-' + dept + '-' + v + '-' + row['Sub Department'] + '-' + row['Office Reporting Location'] + '-' + pid, \
                                                                        allocated_value , ' ', l, '0' + str(deptcodetosub_dict[row['ADP Department Code']]), \
                                                                        'NULL', 'NULL', str(row['COMPANY CODE']) + '- Allocations - PPE ' + row['PERIOD ENDING DATE']]
                    elif (dept in all_alloc_depts) and (cc == 'ML7'):
                        #print("Dept Hit")
                        df_dept_allocations.loc[len(df_dept_allocations.index)] = [ent_template, entitytagging_dict[hc][str(row['Office Reporting Location'])], row['PERIOD ENDING DATE'], row['PAY DATE'], ' ', 'G/L Account', \
                                                                        str(coa_dict[row['Sub Department']][v]), ' ', str(row['COMPANY CODE']) + '-' + str(row['PERIOD ENDING DATE']) + '-' + dept + '-' + v + '-' + row['Sub Department'] + '-' + row['Office Reporting Location'] + '-' + pid, \
                                                                        ' ', row[v], row['Office Reporting Location'], '0' + str(deptcodetosub_dict[row['ADP Department Code']]), \
                                                                        'NULL', 'NULL', str(row['COMPANY CODE']) + '- Allocations - PPE ' + row['PERIOD ENDING DATE']]
                        for l in all_locations:
                            if l == 'SFM MSO':
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
                                df_dept_allocations.loc[len(df_dept_allocations.index)] = [ent_template, entitytagging_dict[str(row['COMPANY CODE'])][l], row['PERIOD ENDING DATE'], row['PAY DATE'], ' ', 'G/L Account', \
                                                                        coa_dict[row['Sub Department']][v], ' ', str(row['COMPANY CODE']) + '-' + str(row['PERIOD ENDING DATE']) + '-' + dept + '-' + v + '-' + row['Sub Department'] + '-' + row['Office Reporting Location'] + '-' + pid, \
                                                                        allocated_value , ' ', l, '0' + str(deptcodetosub_dict[row['ADP Department Code']]), \
                                                                        'NULL', 'NULL', str(row['COMPANY CODE']) + '- Allocations - PPE ' + row['PERIOD ENDING DATE']]
            else:
                missing_headers.append(v)
    mh = set(missing_headers)
    print (" The following headers were missing from the Input file")
    print(mh)

    #print(df_dept_allocations)
    # Start the "Save As" dialog box for the Employee Allocations.
    print("Save the Employee Allocations Output.")                            
    app = tk.Tk()
    app.title("Save File As")
    status_label = tk.Label(app, text="", fg="green")
    status_label.pack()
    save_button = tk.Button(app, text="Save as", command=save_dataframe(df_emp_allocations, status_label))
    save_button.pack(padx=20, pady=10)

    
    # Check if the dept allocations file is empty.
    if len(df_dept_allocations) > 0:
        # Start the "Save As" dialog box for the Employee Allocations
        print("Save the Dept Allocations Output.")                            
        app = tk.Tk()
        app.title("Save File As")
        status_label = tk.Label(app, text="", fg="green")
        status_label.pack()
        save_button = tk.Button(app, text="Save as", command=save_dataframe(df_dept_allocations, status_label))
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
