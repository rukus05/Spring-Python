

import pandas as pd
import time
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
    # If the raw data file has all the data in one file (002, 007, 008 sheets), you must save each entity in it's own file and run the program against each.
    # Read in Data from the "RawData.xlsx" file.
    f = FilePrompt()
    df_spring = pd.read_excel(f)
    

    
    df_spring = df_spring.reset_index()
    
    # Get the unique Invoice Numbers, Locations, Sub Departments, and Department Long Descriptions.  This is needed to loop through and group them properly.
    # uniqueInvoices = df_spring['Invoice Number'].unique()
    unique_Locations = df_spring['LOCATION'].unique()
    unique_SubDept = df_spring['SUB_DEPARTMENT'].unique()
    unique_DptCode = df_spring['DEPT CODE'].unique()
    unique_Entity = df_spring['Entity'].unique()
    
    #  It's important to fill in blank cells for the below columns with Zeros.  A blank cell breaks the calculations
    #  The .fillna() method fills blank cells in these columns with 0's.
    df_spring['ACCRUAL'] = df_spring['ACCRUAL'].fillna(0)
    df_spring['Entity'] = df_spring['Entity'].fillna(0)
    
    
    
    # Create new Dataframe for the Exceptions Output.
    df_exceptions = pd.DataFrame(columns=['Employee Number', 'Last Name', 'First Name', 'SUB_DEPARTMENT', 'DEPT CODE', 'LOCATION', 'Department Long Descr', 'GL Code', 'ACCRUAL'])
    

    

    exc_Dict = {}   # Exclusion Dict
    cc_Dict = {}    # Call Center Dict
    mr_Dict = {}    # Call Center Dict
    rhq_Dict = {}   # Reception HQ Dict    
    fc_Dict = {}    # Financial Counselor Dict
    co_Dict = {}    # Clinical Operations Dict
    SFOAKSV = ['SF', 'OAK', 'SV']
    SFOAKSVNYC = ['SF', 'OAK', 'SV', 'NYC']
    SFOAKSVNYCNEST = ['SF', 'OAK', 'SV', 'NYC', 'Nest']
    HQNEST = ['HQ', 'Nest']
    SFSV = ['SF', 'SV']
    NYCMSO = ['NYC', 'MSO']
    SFHQSV = ['SF', 'HQ', 'SV']


    # Add all Call Center People into a Dict Data Structure
    # Add all Medical Records People into a Dict Data Structure
    for index, row in df_spring.iterrows():
        #if (row['Department Long Descr'] == 'Call Center') and (row['Employee Name'] != 'Lee,Stephannie Victoria'): 
        if re.match('Call Center*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Call Center '):
            cc_Dict[row['Employee Number']] = [0, "", "", "", "", "", "", "", False]
        if re.match('Medical Records*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Medical Records'): 
            mr_Dict[row['Employee Number']] = [0, "", "", "", "", "", "", "", False]
        if re.match('Receptionist HQ*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Receptionist HQ'): 
            rhq_Dict[row['Employee Number']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
        if re.match('Financial Counselor*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Financial Counselor'): 
            fc_Dict[row['Employee Number']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]
        if re.match('Clinical Operations*', str(row['Department Long Descr']), re.IGNORECASE):
        #if (row['Department Long Descr'] == 'Clinical Operations'): 
            co_Dict[row['Employee Number']] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", False]



    for i in unique_SubDept:
        for k in unique_Locations:
            for j in unique_DptCode:
                
                # Phuong Dam
                #exc_Dict[2167864] = [0, "", "", "", "", "", "", "", False]
                # My Dung Lee
                #exc_Dict[1788698] = [0, "", "", "", "", "", "", "", False]
                # Trieu,Minh Hue
                #exc_Dict[1796892] = [0, "", "", "", "", "", "", "", False]
                # Lagano,Lauren
                #exc_Dict[10328183] = [0, "", "", "", "", "", "", "", False]                
                # Bell,Allie
                #exc_Dict[1792388] = [0, "", "", "", "", "", "", "", False]

                #Audrey Krall
                exc_Dict[2495811] = [0, "", "", "", "", "", "", "", False]
                # Dean, Ursula
                exc_Dict[1906920] = [0, "", "", "", "", "", "", "", False]
                # Tran,Nam D
                exc_Dict[1804091] = [0, "", "", "", "", "", "", "", False]
                # Klatsky,Peter
                exc_Dict[1770296] = [0, "", "", "", "", "", "", "", False]
                #Sergio Vaccari
                exc_Dict[1785954] = [0, "", "", "", "", "", "", "", False]
                # Spivey,Allison Michele
                exc_Dict[2566009] = [0, "", "", "", "", "", "", "", False]

                for key in cc_Dict:
                    cc_Dict[key] = [0, "", "", "", "", "", "", "", False]
                for key in mr_Dict:
                    mr_Dict[key] = [0, "", "", "", "", "", "", "", False]
                for key in rhq_Dict:
                    rhq_Dict[key] = [0, "", "", "", "", "", "", "", False]
                for key in fc_Dict:
                    fc_Dict[key] = [0, "", "", "", "", "", "", "", False]
                for key in co_Dict:
                    co_Dict[key] = [0, "", "", "", "", "", "", "", False]

                for index, row in df_spring.iterrows():
                    if (row['SUB_DEPARTMENT'] == i) and (row['DEPT CODE'] == j) and (row['LOCATION'] == k):
                        for key in exc_Dict:
                            if (row['Employee Number'] == key):
                                exc_Dict[key][0] = exc_Dict[key][0] + row['ACCRUAL']
                                exc_Dict[key][1] = row['Last Name']
                                exc_Dict[key][2] = row['First Name']
                                exc_Dict[key][3] = row['SUB_DEPARTMENT']
                                exc_Dict[key][4] = row['DEPT CODE']
                                exc_Dict[key][5] = row['LOCATION']
                                exc_Dict[key][6] = row['Department Long Descr']
                                exc_Dict[key][7] = row['GL Code']
                                exc_Dict[key][8] = True
                                df_spring = df_spring.drop(index)

                        for key2 in cc_Dict:
                            if (row['Employee Number'] == key2):
                                cc_Dict[key2][0] = cc_Dict[key2][0] + row['ACCRUAL']
                                cc_Dict[key2][1] = row['Last Name']
                                cc_Dict[key2][2] = row['First Name']
                                cc_Dict[key2][3] = row['SUB_DEPARTMENT']
                                cc_Dict[key2][4] = row['DEPT CODE']
                                cc_Dict[key2][5] = row['LOCATION']
                                cc_Dict[key2][6] = row['Department Long Descr']
                                cc_Dict[key2][7] = row['GL Code']
                                cc_Dict[key2][8] = True
                                df_spring = df_spring.drop(index)

                        for key3 in mr_Dict:
                            if (row['Employee Number'] == key3):
                                mr_Dict[key3][0] = mr_Dict[key3][0] + row['ACCRUAL']
                                mr_Dict[key3][1] = row['Last Name']
                                mr_Dict[key3][2] = row['First Name']
                                mr_Dict[key3][3] = row['SUB_DEPARTMENT']
                                mr_Dict[key3][4] = row['DEPT CODE']
                                mr_Dict[key3][5] = row['LOCATION']
                                mr_Dict[key3][6] = row['Department Long Descr']
                                mr_Dict[key3][7] = row['GL Code']
                                mr_Dict[key3][8] = True
                                df_spring = df_spring.drop(index)
                        
                        for key4 in rhq_Dict:
                            if (row['Employee Number'] == key4):
                                rhq_Dict[key4][0] = rhq_Dict[key4][0] + row['ACCRUAL']
                                rhq_Dict[key4][1] = row['Last Name']
                                rhq_Dict[key4][2] = row['First Name']
                                rhq_Dict[key4][3] = row['SUB_DEPARTMENT']
                                rhq_Dict[key4][4] = row['DEPT CODE']
                                rhq_Dict[key4][5] = row['LOCATION']
                                rhq_Dict[key4][6] = row['Department Long Descr']
                                rhq_Dict[key4][7] = row['GL Code']
                                rhq_Dict[key4][8] = True
                                df_spring = df_spring.drop(index)

                        for key5 in fc_Dict:
                            if (row['Employee Number'] == key5):
                                fc_Dict[key5][0] = fc_Dict[key5][0] + row['ACCRUAL']
                                fc_Dict[key5][1] = row['Last Name']
                                fc_Dict[key5][2] = row['First Name']
                                fc_Dict[key5][3] = row['SUB_DEPARTMENT']
                                fc_Dict[key5][4] = row['DEPT CODE']
                                fc_Dict[key5][5] = row['LOCATION']
                                fc_Dict[key5][6] = row['Department Long Descr']
                                fc_Dict[key5][7] = row['GL Code']
                                fc_Dict[key5][8] = True
                                df_spring = df_spring.drop(index)

                        for key6 in co_Dict:
                            if (row['Employee Number'] == key6):
                                co_Dict[key6][0] = co_Dict[key6][0] + row['ACCRUAL']
                                co_Dict[key6][1] = row['Last Name']
                                co_Dict[key6][2] = row['First Name']
                                co_Dict[key6][3] = row['SUB_DEPARTMENT']
                                co_Dict[key6][4] = row['DEPT CODE']
                                co_Dict[key6][5] = row['LOCATION']
                                co_Dict[key6][6] = row['Department Long Descr']
                                co_Dict[key6][7] = row['GL Code']
                                co_Dict[key6][8] = True
                                df_spring = df_spring.drop(index)


                for emp in exc_Dict:
                    #if (emp == 'Krall,Audrey') or (emp == 'Dean,Ursula J'):
                    if re.match('2495811*', str(emp)) or re.match('1906920*', str(emp)):
                        if (exc_Dict[emp][8] == True):
                            for loc in SFOAKSV:
                                if loc == 'SF':
                                    pct = 0.5625
                                if loc == 'OAK':
                                    pct = 0.25
                                if loc == 'SV':
                                    pct = 0.1875
                                alloc_ACCRUAL = exc_Dict[emp][0] * pct
                                ln = exc_Dict[emp][1]
                                fn = exc_Dict[emp][2]
                                sd = exc_Dict[emp][3]
                                dc = exc_Dict[emp][4]
                                locn = loc
                                dld = exc_Dict[emp][6]
                                gl = exc_Dict[emp][7]

                                df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                    '''
                    ### These 2 if statements are now covered by co_Dict###

                    #if re.match('Dam,Phuong*', str(emp), re.IGNORECASE) or re.match('Trieu,Minh*', str(emp), re.IGNORECASE) or re.match('Bell,Allie*', str(emp), re.IGNORECASE):
                    if re.match('2167864*', str(emp)) or re.match('1796892*', str(emp)) or re.match('1792388*', str(emp)):
                        if (exc_Dict[emp][8] == True):
                            for loc in SFOAKSVNYC:
                                pct = 0.25
                                alloc_ACCRUAL = exc_Dict[emp][0] * pct
                                ln = exc_Dict[emp][1]
                                fn = exc_Dict[emp][2]
                                sd = exc_Dict[emp][3]
                                dc = exc_Dict[emp][4]
                                locn = loc
                                dld = exc_Dict[emp][6]
                                gl = exc_Dict[emp][7]
                                
                                df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                            
                    #if re.match('Lee,My Dung*', str(emp), re.IGNORECASE) and (exc_Dict[emp][16] == True):
                    if re.match('1788698*', str(emp)) and (exc_Dict[emp][8] == True):
                        for loc in SFOAKSVNYCNEST:
                            if loc == 'Nest':
                                pct = 0.1
                            else:
                                pct = 0.225
                            alloc_ACCRUAL = exc_Dict[emp][0] * pct
                            ln = exc_Dict[emp][1]
                            fn = exc_Dict[emp][2]
                            sd = exc_Dict[emp][3]
                            dc = exc_Dict[emp][4]
                            locn = loc
                            dld = exc_Dict[emp][6]
                            gl = exc_Dict[emp][7]

                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                    '''
                    #if re.match('Tran,Nam D*', str(emp), re.IGNORECASE) and (exc_Dict[emp][16] == True):
                    if re.match('1804091*', str(emp)) and (exc_Dict[emp][8] == True):
                        
                        for loc in SFHQSV:
                            if loc == 'SF':
                                pct = 0.59
                            if loc == 'HQ':
                                pct = 0.1
                            if loc == 'SV':
                                pct = 0.31
                            alloc_ACCRUAL = exc_Dict[emp][0] * pct
                            ln = exc_Dict[emp][1]
                            fn = exc_Dict[emp][2]
                            sd = exc_Dict[emp][3]
                            dc = exc_Dict[emp][4]
                            locn = loc
                            #locn = exc_Dict[emp][5]
                            dld = exc_Dict[emp][6]
                            gl = exc_Dict[emp][7]

                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]

                    
                    #if re.match('Klatsky,Peter*', str(emp), re.IGNORECASE) and (exc_Dict[emp][16] == True):
                    if re.match('1770296*', str(emp)) and (exc_Dict[emp][8] == True):
                        
                        for loc in NYCMSO:
                            if loc == 'NYC':
                                pct = 0.75
                            else:
                                pct = 0.25
                            alloc_ACCRUAL = exc_Dict[emp][0] * pct
                            ln = exc_Dict[emp][1]
                            fn = exc_Dict[emp][2]
                            sd = exc_Dict[emp][3]
                            dc = exc_Dict[emp][4]
                            locn = loc
                            dld = exc_Dict[emp][6]
                            gl = exc_Dict[emp][7]

                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                    
                    #  Sergio Vaccari
                    if re.match('1785954*', str(emp)) and (exc_Dict[emp][8] == True):
                    #if (emp == 'Vaccari,Sergio') and (exc_Dict[emp][16] == True):
                        #  Loop Through Sergio Vaccari Locations in Exception List
                        for loc in SFOAKSV:
                            if loc == 'SF':
                                pct = 0.5625
                            if loc == 'OAK':
                                pct = 0.25
                            if loc == 'SV':    
                                pct = 0.1875
                            alloc_ACCRUAL = exc_Dict[emp][0] * pct
                            ln = exc_Dict[emp][1]
                            fn = exc_Dict[emp][2]
                            sd = exc_Dict[emp][3]
                            dc = exc_Dict[emp][4]
                            locn = loc
                            dld = exc_Dict[emp][6]
                            gl = exc_Dict[emp][7]

                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]

                    #  Spivey,Allison Michele
                    if re.match('2566009*', str(emp)) and (exc_Dict[emp][8] == True):
                    #if (emp == 'Vaccari,Sergio') and (exc_Dict[emp][16] == True):
                        #  Loop Through Sergio Vaccari Locations in Exception List
                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':    
                                pct = 0.15
                            if loc == 'NYC':    
                                pct = 0.2
                            alloc_ACCRUAL = exc_Dict[emp][0] * pct
                            ln = exc_Dict[emp][1]
                            fn = exc_Dict[emp][2]
                            sd = exc_Dict[emp][3]
                            dc = exc_Dict[emp][4]
                            locn = loc
                            dld = exc_Dict[emp][6]
                            gl = exc_Dict[emp][7]

                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                    '''
                    #if re.match('Legano,Lauren*', str(emp), re.IGNORECASE) and (exc_Dict[emp][16] == True):
                    if re.match('10328183*', str(emp)) and (exc_Dict[emp][8] == True):
                        #  Loop Through Sergio Vaccari Locations in Exception List
                        for loc in HQNEST:
                            pct = 0.5
                            alloc_ACCRUAL = exc_Dict[emp][0] * pct
                            ln = exc_Dict[emp][1]
                            fn = exc_Dict[emp][2]
                            sd = exc_Dict[emp][3]
                            dc = exc_Dict[emp][4]
                            locn = loc
                            dld = exc_Dict[emp][6]
                            gl = exc_Dict[emp][7]

                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                    '''
                
                for emp in cc_Dict:
                    #  
                    if cc_Dict[emp][8] == True:
                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':    
                                pct = 0.15
                            if loc == 'NYC':
                                pct = 0.2
                            alloc_ACCRUAL = cc_Dict[emp][0] * pct
                            ln = cc_Dict[emp][1]
                            fn = cc_Dict[emp][2]
                            sd = cc_Dict[emp][3]
                            dc = cc_Dict[emp][4]
                            locn = loc
                            dld = cc_Dict[emp][6]
                            gl = cc_Dict[emp][7]
                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]

                for emp in mr_Dict:
                    #  
                    if mr_Dict[emp][8] == True:

                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':    
                                pct = 0.15
                            if loc == 'NYC':
                                pct = 0.2
                            alloc_ACCRUAL = mr_Dict[emp][0] * pct
                            ln = mr_Dict[emp][1]
                            fn = mr_Dict[emp][2]
                            sd = mr_Dict[emp][3]
                            dc = mr_Dict[emp][4]
                            locn = loc
                            dld = mr_Dict[emp][6]
                            gl = mr_Dict[emp][7]
                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]

                for emp in rhq_Dict:
                    #  
                    if rhq_Dict[emp][8] == True:

                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':    
                                pct = 0.15
                            if loc == 'NYC':
                                pct = 0.2
                            alloc_ACCRUAL = rhq_Dict[emp][0] * pct
                            ln = rhq_Dict[emp][1]
                            fn = rhq_Dict[emp][2]
                            sd = rhq_Dict[emp][3]
                            dc = rhq_Dict[emp][4]
                            locn = loc
                            dld = rhq_Dict[emp][6]
                            gl = rhq_Dict[emp][7]
                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                
                for emp in fc_Dict:
                    #  
                    if fc_Dict[emp][8] == True:

                        for loc in SFOAKSVNYC:
                            if loc == 'SF':
                                pct = 0.45
                            if loc == 'OAK':
                                pct = 0.2
                            if loc == 'SV':    
                                pct = 0.15
                            if loc == 'NYC':
                                pct = 0.2
                            alloc_ACCRUAL = fc_Dict[emp][0] * pct
                            ln = fc_Dict[emp][1]
                            fn = fc_Dict[emp][2]
                            sd = fc_Dict[emp][3]
                            dc = fc_Dict[emp][4]
                            locn = loc
                            dld = fc_Dict[emp][6]
                            gl = fc_Dict[emp][7]
                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]

                for emp in co_Dict:
                    #  My Dung Lee
                    if re.match('1788698*', str(emp)) and (co_Dict[emp][8] == True):
                        for loc in SFOAKSVNYCNEST:
                            if loc == 'Nest':
                                pct = 0.1
                            if loc == 'SF':
                                pct = 0.405
                            if loc == 'OAK':
                                pct = 0.18
                            if loc == 'SV':
                                pct = 0.135
                            if loc == 'NYC':
                                pct = 0.18
                            alloc_ACCRUAL = co_Dict[emp][0] * pct
                            ln = co_Dict[emp][1]
                            fn = co_Dict[emp][2]
                            sd = co_Dict[emp][3]
                            dc = co_Dict[emp][4]
                            locn = loc
                            dld = co_Dict[emp][6]
                            gl = co_Dict[emp][7]
                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]

                    #  Lauren Lagano
                    elif re.match('10328183*', str(emp)) and (co_Dict[emp][8] == True):
                        for loc in HQNEST:
                            pct = 0.5
                            alloc_ACCRUAL = co_Dict[emp][0] * pct
                            ln = co_Dict[emp][1]
                            fn = co_Dict[emp][2]
                            sd = co_Dict[emp][3]
                            dc = co_Dict[emp][4]
                            locn = loc
                            dld = co_Dict[emp][6]
                            gl = co_Dict[emp][7]
                            df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                    
                    else:
                        if co_Dict[emp][8] == True:
                             for loc in SFOAKSVNYC:
                                if loc == 'SF':
                                    pct = 0.45
                                if loc == 'OAK':
                                    pct = 0.2
                                if loc == 'SV':
                                    pct = 0.15
                                if loc == 'NYC':
                                    pct = 0.2
                                alloc_ACCRUAL = co_Dict[emp][0] * pct
                                ln = co_Dict[emp][1]
                                fn = co_Dict[emp][2]
                                sd = co_Dict[emp][3]
                                dc = co_Dict[emp][4]
                                locn = loc
                                dld = co_Dict[emp][6]
                                gl = co_Dict[emp][7]
                                df_exceptions.loc[len(df_exceptions.index)] = [emp, ln, fn, sd, dc, locn, dld, gl, alloc_ACCRUAL]
                    




    df_concatenated = pd.concat([df_spring, df_exceptions], ignore_index=True).fillna(0)
    
    df_concatenated.reset_index()
    
    #df_group = df_concatenated.groupby(['SUB_DEPARTMENT', 'LOCATION', 'DEPT CODE', ])
    
    
    unique_Locations2 = df_concatenated['LOCATION'].unique()
    unique_SubDept2 = df_concatenated['SUB_DEPARTMENT'].unique()
    unique_DptCode2 = df_concatenated['DEPT CODE'].unique()
    unique_Entity2 = df_concatenated['Entity'].unique()


    # Create new Dataframe for the Output.
    df_Output = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', 'DebitAmt', 'CreditAmt', 'Loc', 'Dept', 'Provider', 'Service Line', 'Comments'])
    
    # By using 4 nested FOR loops, we can group the rows by Invoices, Dept Long Desc, Sup Dept, and Location
    for i in unique_SubDept2:
        for j in unique_Locations2:
            for k in unique_DptCode2:
                # Initialize Sum Variables
                Accrual_Sum = 0
                                
                # Loop through rows of the Raw Data file.
                for index, row in df_concatenated.iterrows():
                    # These if statements force a match for specific Invoices, Dept Long Desc, Sup Dept, and Locations
                    if (row['SUB_DEPARTMENT'] == i) and (row['LOCATION'] == j) and (row['DEPT CODE'] == k):
                        # Sum up the pertinent columns.
                        Accrual_Sum = Accrual_Sum + row['ACCRUAL']
                        glCode = row['GL Code']
                        deptCode = row['DEPT CODE']
                        comment = row['Department Long Descr']
                        #subDept = row['SUB_DEPARTMENT']
                       
                       

                # Create an array holding the Sums for a particular Invoice, DLD, Location, and Sub_Dept        
                
                #summary = [i, j, k , Accrual_Sum]          
                                
                
                # Initialize counter for CoA
                #cnt = 4
                # Loop over the Sum variables in the Summary Array.  
                # Add row to output file if GrossWages doesn't equal 0.  
                # Each Row in the loop will be a debit entry for a particular sum variable (as defined above)
                
                if Accrual_Sum != 0:  
                    df_Output.loc[len(df_Output.index)] = ["", "", "", "", i, glCode, "", "", Accrual_Sum, "", j, deptCode, "", "", comment]
                #   cnt = cnt + 1       # Increment CoA counter
                    
                # Add a credit entry to match the above debits.
                #if GrossWages_Sum != 0:     
                #    df_Output.loc[len(df_Output.index)] = [ "", ped.date(), ivd.date(), "", str(k), 23300, "", str(i) + ' ' + str(j), "", TotClientCharges_Sum , l, deptCode, "", "", ""]
                
                

    #  Export the Pandas DataFrame to an Excel file called "Output_test.xlsx"
    inp = input("Please type name of file for Output:")
    des = str(inp + '.xlsx')
    df_Output.to_excel(des, index = False)

    runningtime = time.time() - start
    print("The time for this script is:", runningtime)

    

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



