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
from copy import deepcopy
from zpack.fns import FilePrompt
from zpack.fns import save_dataframe



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
    #print(dept_alloc_dict)
    
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
    # Entity dict has the following format:
    # {'ML7': {'SFM MSO': '002', 'Nest': '002', 'SF': '010', 'OAK': '011', 'SV': '012', 'NYC': '013', 'PDX': '014'}, 
    # '22J': {'SFM MSO': '002', 'Nest': '002', 'SF': '007', 'OAK': '007', 'SV': '007', 'NYC': '013', 'PDX': '014'}, 
    # '362': {'SFM MSO': '002', 'Nest': '002', 'SF': '010', 'OAK': '011', 'SV': '012', 'NYC': '008', 'PDX': '014'}, 
    # '633': {'SFM MSO': '002', 'Nest': '002', 'SF': '010', 'OAK': '011', 'SV': '012', 'NYC': '013', 'PDX': '009'}}
    

    # Prompt user for Chart of Accounts File
    print("Select the Chart of Accounts File:")
    coaf = FilePrompt()
    coa_df = pd.read_excel(coaf)
    coa_df = coa_df.reset_index()
    coa_dict = chart_of_accounts(coa_df)

    # coa_dict has the following format:
    #{'ASC': {'index': 0, 'SUB_DEPARTMENT': 'ASC', 'Salaries and Wages': 51111, 'OT': 51121, 'ELC': 65200, 'ER Taxes': 51141, '401K-ER Match': 51161, 'Medical Waiver': 51171, 'MEDICAL': 51171, 'DENTAL': 51171, 'VISION': 51171, 'LIFE': 51171, 'Other Benefits': 51171}, 
    # 'Clinical': {'index': 1, 'SUB_DEPARTMENT': 'Clinical', 'Salaries and Wages': 51110, 'OT': 51120, 'ELC': 65200, 'ER Taxes': 51140, '401K-ER Match': 51160, 'Medical Waiver': 51170, 'MEDICAL': 51170, 'DENTAL': 51170, 'VISION': 51170, 'LIFE': 51170, 'Other Benefits': 51170}, 
    # 'HQ': {'index': 2, 'SUB_DEPARTMENT': 'HQ', 'Salaries and Wages': 61110, 'OT': 61120, 'ELC': 65200, 'ER Taxes': 61140, '401K-ER Match': 61170, 'Medical Waiver': 61180, 'MEDICAL': 61180, 'DENTAL': 61180, 'VISION': 61180, 'LIFE': 61180, 'Other Benefits': 61180}, 
    # 'MD': {'index': 3, 'SUB_DEPARTMENT': 'MD', 'Salaries and Wages': 51113, 'OT': 51123, 'ELC': 65200, 'ER Taxes': 51143, '401K-ER Match': 51163, 'Medical Waiver': 51173, 'MEDICAL': 51173, 'DENTAL': 51173, 'VISION': 51173, 'LIFE': 51173, 'Other Benefits': 51173}, 
    # 'Lab': {'index': 4, 'SUB_DEPARTMENT': 'Lab', 'Salaries and Wages': 51112, 'OT': 51122, 'ELC': 65200, 'ER Taxes': 51142, '401K-ER Match': 51162, 'Medical Waiver': 51172, 'MEDICAL': 51172, 'DENTAL': 51172, 'VISION': 51172, 'LIFE': 51172, 'Other Benefits': 51172}, 
    # 'Operating': {'index': 5, 'SUB_DEPARTMENT': 'Operating', 'Salaries and Wages': 61110, 'OT': 61120, 'ELC': 65200, 'ER Taxes': 61140, '401K-ER Match': 61170, 'Medical Waiver': 61180, 'MEDICAL': 61180, 'DENTAL': 61180, 'VISION': 61180, 'LIFE': 61180, 'Other Benefits': 61180}}

    

    # Prompt user for Input file
    print("Select the Input File which your running Payroll Allocations for:")
    inputf = FilePrompt()
    df = pd.read_excel(inputf,dtype={'FILE NUMBER': str})
    df = df.reset_index()
    #print(df)

    # Fill in blank cells with 0
    df = df.fillna(0)
    # Obtain the Company Code, Period Ending Date, and Pay Date of the input file.
    company_code = str(df.at[0,'COMPANY CODE'])
    company_code = company_code.rstrip('.0')
    print('The company code is :', company_code)
    ped = df.at[0, 'PERIOD ENDING DATE']
    payd = df.at[0, 'PAY DATE']
    df['ADP Department Code'] = df['ADP Department Code'].astype(int)
    
    # Create new Dataframe for the Employee Allocations Output.
    df_emp_allocations = pd.DataFrame(columns=['Entity Template', 'Entity', 'PostDate', 'DocDate', 'DocNo','AcctType', 'AcctNo', 'AcctName', 'Description', \
                                           'DebitAmt', 'CreditAmt', 'Loc', 'Dept','Provider', 'Service Line', 'Comments'])
    # Create new Dataframe for the Dept Allocations Output.
    df_dept_allocations = pd.DataFrame(columns=['Entity Template', 'Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', \
                                           'DebitAmt', 'CreditAmt', 'Loc', 'Dept','Provider', 'Service Line', 'Comments'])
    
    # Capture home payroll company of the file.
    if (company_code  == '362') or (company_code  == '362.0'):
        hc = '362'
        ent_template = '008'
    elif company_code  == '22J':
    # elif str(df.loc[1, 'COMPANY CODE'])  == '22J':
        hc = '22J'
        ent_template = '007'
    elif company_code == 'ML7':
        hc = 'ML7'
        ent_template = '002'
    
    print(hc)
    # Create a list of all Locations
    all_locations = ['SFM MSO', 'Nest', 'SF', 'OAK', 'SV', 'NYC', 'PDX']
    # Create list for ALlocated depts based on the Dept Allocation Dictionary, which was created from the allocations file.
    all_alloc_depts = list(dept_alloc_dict.keys())
    
    #all_alloc_depts = ['Receptionist HQ', 'Medical Records', 'Call Center', 'Financial Counselor', 'Clinical Operations', 'Revenue Cycle']
    
    #l_dict = {'SFM MSO' : 0, 'Nest' : 0, 'SF' : 0, 'OAK' : 0, 'SV' : 0, 'NYC' : 0, 'PDX' : 0}
    
    #dept_dict_alloc_values = {dept : deepcopy(l_dict) for dept in all_alloc_depts}
    
    
    # Create a list of all values to allocate
    coa_headers = coa_df.columns
    all_values = coa_headers.tolist()
    
    # Initialize the sums of 'All Value' items to 0.
    
    dict_allValues_sum = {key : 0 for key in all_values[2:]}
    # Here's what the dict looks like:
    # {'SUB_DEPARTMENT': 0, 'Salaries and Wages': 0, 'OT': 0, 'ELC': 0, 'ER Taxes': 0, '401K-ER Match': 0, 
    # 'Medical Waiver': 0, 'MEDICAL': 0, 'DENTAL': 0, 'VISION': 0, 'LIFE': 0, 'Other Benefits': 0}
    
    
    # Create a dict that assigns each dept to a value that is a copy of the above  initialized dict_allValues_sum dictionary.
    dict_usefor_sumV = {dept : deepcopy(dict_allValues_sum) for dept in all_alloc_depts}
    # Here's what the dict looks like:
    # {'Receptionist HQ': {'Salaries and Wages': 0, 'OT': 0, 'ELC': 0, 'ER Taxes': 0, '401K-ER Match': 0, 'Medical Waiver': 0, 'MEDICAL': 0, 'DENTAL': 0, 'VISION': 0, 'LIFE': 0, 'Other Benefits': 0}, 
    # {'Medical Records': {'Salaries and Wages': 0, 'OT': 0, 'ELC': 0, 'ER Taxes': 0, '401K-ER Match': 0, 'Medical Waiver': 0, 'MEDICAL': 0, 'DENTAL': 0, 'VISION': 0, 'LIFE': 0, 'Other Benefits': 0}, 
    # etc, etc

    # Create a dict that has something like:  {'Nest' : 0, 'OAK' : 0, 'NYC' : 0, etc}
    # Where the values will be the percentage allocations; initialized to 0.
    dict_loc_to_pctv = {loc : 0 for loc in all_locations}

    # Create a dict that assigned each Value that's being allocated for to a clone of the percentage allocations dict, dict_loc_to_pctv
    dict_allValues_to_locpctv = {allv : deepcopy(dict_loc_to_pctv) for allv in all_values[2:]}
    # The dictionary looks like:
    # {'Salaries and Wages': {'Nest' : 0, 'OAK' : 0, 'NYC' : 0, etc}, 'OT': {'Nest' : 0, 'OAK' : 0, 'NYC' : 0, etc}, etc, etc}
    # The 2: skips 'Index' and 'SUB_DEPARTMENT'; start at element #2

    # Now, crate a dict that assigns each dept to a clone of the dict_allValues_to_locpctv dict:
    dict_usefor_pctV = {dept : deepcopy(dict_allValues_to_locpctv) for dept in all_alloc_depts }
    # The dict looks like: 
    # {'Call Center' : {'Salaries and Wages': {'Nest' : 0, 'OAK' : 0, 'NYC' : 0, etc}, 'OT': {'Nest' : 0, 'OAK' : 0, 'NYC' : 0, etc}, etc, etc},
    # {'Medical Records' : {'Salaries and Wages': {'Nest' : 0, 'OAK' : 0, 'NYC' : 0, etc}, 'OT': {'Nest' : 0, 'OAK' : 0, 'NYC' : 0, etc}, etc, etc},

    
    # Here's what the final dict for the allocated values, dict_usefor_pctV, will look like:
    # {'Receptionist HQ': {'Salaries and Wages': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'OT': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ELC': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ER Taxes': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, '401K-ER Match': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Medical Waiver': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'MEDICAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'DENTAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'VISION': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'LIFE': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Other Benefits': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}}, 
    # 'Medical Records': {'Salaries and Wages': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'OT': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ELC': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ER Taxes': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, '401K-ER Match': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Medical Waiver': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'MEDICAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'DENTAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'VISION': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'LIFE': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Other Benefits': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}}, 
    # 'Call Center': {'Salaries and Wages': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'OT': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ELC': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ER Taxes': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, '401K-ER Match': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Medical Waiver': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'MEDICAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'DENTAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'VISION': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'LIFE': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Other Benefits': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}}, 
    # 'Financial Counselor': {'Salaries and Wages': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'OT': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ELC': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ER Taxes': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, '401K-ER Match': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Medical Waiver': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'MEDICAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'DENTAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'VISION': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'LIFE': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Other Benefits': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}}, 
    # 'Clinical Operations': {'Salaries and Wages': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'OT': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ELC': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ER Taxes': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, '401K-ER Match': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Medical Waiver': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'MEDICAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'DENTAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'VISION': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'LIFE': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Other Benefits': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}}, 
    # 'Revenue Cycle': {'Salaries and Wages': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'OT': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ELC': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'ER Taxes': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, '401K-ER Match': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Medical Waiver': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'MEDICAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'DENTAL': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'VISION': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'LIFE': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}, 'Other Benefits': {'SFM MSO': 0, 'Nest': 0, 'SF': 0, 'OAK': 0, 'SV': 0, 'NYC': 0, 'PDX': 0}}}

    # Create a dict to associate Depts with Sub Depts
    dict_dept_to_subdept = {dept : '' for dept in all_alloc_depts}
    
    # Create a dict for Dept to ADP Code mappings
    dict_dept_to_ADPCode = {dept : '' for dept in all_alloc_depts}

    # Create a list of the allocations file headers.  
    alloc_headers = df.columns
    alloc_headers_values = alloc_headers.tolist()
    #  Below lines are attempts to set all headers to lower case for matching CoA.
    # ahv_lower = [s.lower() for s in alloc_headers_values]
    # print(alloc_headers_values)


    # Create a set to capture Allocation headers that are not in the input file
    missing_headers = []
    # Initialize agg_v to zero.  Used to check if Dept allocations are used for this file.
    agg_v = 0

    
    for index, row in df.iterrows():
        
        pid = str(row['POSITION ID'])
        print(pid)
        
        
        dept = row['Department']
        cc = str(row['COMPANY CODE'])
        cc = cc.rstrip('.0')
        print('The cc reread is : ', cc)
        

        # For some reason, the 362 files add a ".0" at the end.  Hence, we're stripping it away for 362 files.
        if (cc == '362') or (cc == '362.0'):
            pid = pid.rstrip('.0')
            print('The pid is: ', pid)
        #print('The cc is:', cc)
        #print(type(cc))
        
        #if (cc == '362') or (cc == '362.0'):
        #    print('hit')
        #    cc = '362'
        #    print('The pid is: ', pid)
        '''
        Legacy Code
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
        '''
        # Check if dept is Receptionist HQ. If so, check if there is a seperate allocation for the employee. (2nd if)
        # If so, use the % as defined in the dictionary when create_empalloc_dict is called.
        # If not, use the % as defined in the dictionary when create_deptalloc_dict is called.
        # This process repeats for other depts.  If we don't match any of these, use the emp allocation; last else statement below.
        if re.search('Receptionist HQ*', str(dept), re.IGNORECASE):
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

        elif re.search('Financial Counselor*', str(dept), re.IGNORECASE):
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
        
        #  df_allocations = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'AcctType', 'AcctNo', 'AcctName', 'Description', \
        #                                        'DebitAmt', 'CreditAmt', 'Loc', 'Dept','Provider', 'Service Line', 'Comments'])

        for v in all_values[2:]:
            # This commented out item bypassed checking for all the items in the COA file.
            if v in alloc_headers_values:
                if row[v] != 0.0:
                    if pid in emp_alloc_dict:  # if employee has allocation, use that allocation.  if not, use department allocatioin 
                        # If 'NEST' is an office reporting location, change to 'Nest'
                        s = str(row['Office Reporting Location'])
                        if s == 'NEST':
                            lower_est = s[-3:]
                            corrected_s = 'N' + lower_est.lower()
                        else:
                            corrected_s = s
                        df_emp_allocations.loc[len(df_emp_allocations.index)] = [ent_template, entitytagging_dict[hc][corrected_s], row['PERIOD ENDING DATE'], row['PAY DATE'], ' ', 'G/L Account', \
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
                            
                            print(pct)
                            #if pct == 0.0:
                                #print('Zero')
                            if pct != 0.0:
                                #print('Not Zero')
                                allocated_value = row[v]*pct
                                df_emp_allocations.loc[len(df_emp_allocations.index)] = [ent_template, entitytagging_dict[cc][l], row['PERIOD ENDING DATE'], row['PAY DATE'], ' ', 'G/L Account', \
                                                                        coa_dict[row['Sub Department']][v], ' ', str(row['COMPANY CODE']) + '-' + str(row['PERIOD ENDING DATE']) + '-' + dept + '-' + v + '-' + row['Sub Department'] + '-' + row['Office Reporting Location'] + '-' + pid, \
                                                                        allocated_value , ' ', l, '0' + str(deptcodetosub_dict[row['ADP Department Code']]), \
                                                                        'NULL', 'NULL', str(row['COMPANY CODE']) + '- Allocations - PPE ' + row['PERIOD ENDING DATE']]
                    elif (dept in all_alloc_depts) and (cc == 'ML7'):
                        '''
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
                        '''
                        agg_v = row[v]
                        dict_usefor_sumV[dept][v] = dict_usefor_sumV[dept][v] + agg_v
                        dict_dept_to_subdept[dept] = row['Sub Department']
                        dict_dept_to_ADPCode[dept] = row['ADP Department Code']
                        off_report_loc = row['Office Reporting Location']
                        


            else:
                missing_headers.append(v)
    mh = set(missing_headers)
    print (" The following headers were missing from the Input file")
    print(mh)
    #print(dict_usefor_sumV)
    if agg_v != 0:
        for d, v in dict_usefor_sumV.items():
            hq_percent = dept_alloc_dict[d]['SFM MSO']
            nest_percent = dept_alloc_dict[d]['Nest']
            sf_percent = dept_alloc_dict[d]['SF']
            oak_percent = dept_alloc_dict[d]['OAK']
            sv_percent = dept_alloc_dict[d]['SV']
            nyc_percent = dept_alloc_dict[d]['NYC']
            pdx_percent = dept_alloc_dict[d]['PDX']

            for vals in all_values[2:]:
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
                    
                    dict_usefor_pctV[d][vals][l] = dict_usefor_sumV[d][vals]*pct

        #print(dict_usefor_pctV)

        for outer_key, d_a_lpct in dict_usefor_pctV.items():
            #print(outer_key)
            for vs, locpcts in d_a_lpct.items():
                
                df_dept_allocations.loc[len(df_dept_allocations.index)] = [ent_template, entitytagging_dict[company_code]['SFM MSO'] , ped, payd, ' ', 'G/L Account', coa_dict[dict_dept_to_subdept[outer_key]][vs], ' ', company_code + '-' + ped + '-' + outer_key + '-' + dict_dept_to_subdept[outer_key] + '-' + vs + '-' + off_report_loc, ' ', \
                                                                        dict_usefor_sumV[outer_key][vs], 'SFM MSO', str(deptcodetosub_dict[dict_dept_to_ADPCode[outer_key]]).zfill(4), 'NULL', 'NULL', company_code + '- Payroll Allocations - PPE ' + ped]
                
                for locs, pctvs in locpcts.items():
                        df_dept_allocations.loc[len(df_dept_allocations.index)] = [ent_template, entitytagging_dict[company_code][locs], ped, payd, ' ', 'G/L Account', coa_dict[dict_dept_to_subdept[outer_key]][vs], ' ', company_code + '-' + ped + '-' + outer_key + '-' + dict_dept_to_subdept[outer_key] + '-' + vs + '-' + off_report_loc, \
                                                                            dict_usefor_pctV[outer_key][vs][locs], ' ', locs, str(deptcodetosub_dict[dict_dept_to_ADPCode[outer_key]]).zfill(4), 'NULL', 'NULL', company_code + '- Payroll Allocations - PPE ' + ped]





    #print(df_emp_allocations)
    #print(df_dept_allocations)
    # Start the "Save As" dialog box for the Employee Allocations.
    runningtime = time.time() - start
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
    
    print("The execution time is:", runningtime)



if __name__ == "__main__":
    main()
