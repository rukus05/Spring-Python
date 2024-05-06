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

def chart_of_accounts(input_df):
    
    coa_dict = {}

    for index, row in input_df.iterrows():
        row_dict = row.to_dict()
        coa_dict[row['SUB_DEPARTMENT']] = row_dict
    
    return coa_dict

