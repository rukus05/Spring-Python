import pandas as pd



## This python file creates the dictionaries that contains all allocaiton information


# This function is used to create the dictionary for Employee Allocations
def create_empalloc_dict(input_df): 
    
    # Create the empty dictionary
    ea_dict = {}
    # Iterate through the DataFrame and get location percentages by employee.
    for index, row in input_df.iterrows():
        # 'ALL' is used to designate the department allocations, addressed below.
        if row['Employee Name'] != 'ALL':
            ea_dict[row['POSITION ID']] = {
                'SFM MSO': row['SFM MSO'],
                'Nest': row['Nest'],
                'SF': row['SF'],
                'OAK': row['OAK'],
                'SV': row['SV'],
                'NYC': row['NYC'],
                'PDX': row['PDX']
            
        }
    return ea_dict

# This function is used to create the dictionary for Department Allocations
def create_deptalloc_dict(input_df): 
    # Create the empty dictionary
    dept_dict = {}
    
    for index, row in input_df.iterrows():
        if row['Employee Name'] == 'ALL':
            dept_dict[row['Department Long Descr']] = {
                'SFM MSO': row['SFM MSO'],
                'Nest': row['Nest'],
                'SF': row['SF'],
                'OAK': row['OAK'],
                'SV': row['SV'],
                'NYC': row['NYC'],
                'PDX': row['PDX']
            
        }
    return dept_dict

# This function is used to create the dictionary that maps ADP Dept Codes to Spring Sub Depts
def deptcode_to_subdept(input_df):
    columns = input_df.columns
    dts_dict = {}
    
        
    # Create the dictionary.  'Home Department' : 'Department Code'
    for index, row in input_df.iterrows():
        dts_dict[row['HOME DEPARTMENT']] = int(row['Department Code'])
    
    return dts_dict


 # Entity dict has the following format:
# {'ML7': {'SFM MSO': '002', 'Nest': '002', 'SF': '010', 'OAK': '011', 'SV': '012', 'NYC': '013', 'PDX': '014'}, 
# '22J': {'SFM MSO': '002', 'Nest': '002', 'SF': '007', 'OAK': '007', 'SV': '007', 'NYC': '013', 'PDX': '014'}, 
# '362': {'SFM MSO': '002', 'Nest': '002', 'SF': '010', 'OAK': '011', 'SV': '012', 'NYC': '008', 'PDX': '014'}, 
# '633': {'SFM MSO': '002', 'Nest': '002', 'SF': '010', 'OAK': '011', 'SV': '012', 'NYC': '013', 'PDX': '009'}}
def entity_tagging(input_df):
    entity_tag_dict = {}
    

    for index, row in input_df.iterrows():
        entity_tag_dict[row['Company Code']] = {
            'SFM MSO': row['SFM MSO'],
            'Nest': row['Nest'],
            'SF': row['SF'],
            'OAK': row['OAK'],
            'SV': row['SV'],
            'NYC': row['NYC'],
            'PDX': row['PDX']
        }
    return entity_tag_dict
# coa_dict has the following format:
#{'ASC': {'index': 0, 'SUB_DEPARTMENT': 'ASC', 'Salaries and Wages': 51111, 'OT': 51121, 'ELC': 65200, 'ER Taxes': 51141, '401K-ER Match': 51161, 'Medical Waiver': 51171, 'MEDICAL': 51171, 'DENTAL': 51171, 'VISION': 51171, 'LIFE': 51171, 'Other Benefits': 51171}, 
# 'Clinical': {'index': 1, 'SUB_DEPARTMENT': 'Clinical', 'Salaries and Wages': 51110, 'OT': 51120, 'ELC': 65200, 'ER Taxes': 51140, '401K-ER Match': 51160, 'Medical Waiver': 51170, 'MEDICAL': 51170, 'DENTAL': 51170, 'VISION': 51170, 'LIFE': 51170, 'Other Benefits': 51170}, 
# 'HQ': {'index': 2, 'SUB_DEPARTMENT': 'HQ', 'Salaries and Wages': 61110, 'OT': 61120, 'ELC': 65200, 'ER Taxes': 61140, '401K-ER Match': 61170, 'Medical Waiver': 61180, 'MEDICAL': 61180, 'DENTAL': 61180, 'VISION': 61180, 'LIFE': 61180, 'Other Benefits': 61180}, 
# 'MD': {'index': 3, 'SUB_DEPARTMENT': 'MD', 'Salaries and Wages': 51113, 'OT': 51123, 'ELC': 65200, 'ER Taxes': 51143, '401K-ER Match': 51163, 'Medical Waiver': 51173, 'MEDICAL': 51173, 'DENTAL': 51173, 'VISION': 51173, 'LIFE': 51173, 'Other Benefits': 51173}, 
# 'Lab': {'index': 4, 'SUB_DEPARTMENT': 'Lab', 'Salaries and Wages': 51112, 'OT': 51122, 'ELC': 65200, 'ER Taxes': 51142, '401K-ER Match': 51162, 'Medical Waiver': 51172, 'MEDICAL': 51172, 'DENTAL': 51172, 'VISION': 51172, 'LIFE': 51172, 'Other Benefits': 51172}, 
# 'Operating': {'index': 5, 'SUB_DEPARTMENT': 'Operating', 'Salaries and Wages': 61110, 'OT': 61120, 'ELC': 65200, 'ER Taxes': 61140, '401K-ER Match': 61170, 'Medical Waiver': 61180, 'MEDICAL': 61180, 'DENTAL': 61180, 'VISION': 61180, 'LIFE': 61180, 'Other Benefits': 61180}}

def chart_of_accounts(input_df):
    
    coa_dict = {}

    for index, row in input_df.iterrows():
        row_dict = row.to_dict()
        coa_dict[row['SUB_DEPARTMENT']] = row_dict
    
    return coa_dict

