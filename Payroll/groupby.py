import pandas as pd
import re


#def main(): 
df = pd.read_excel("002-allJuly-alloc.xlsx")
#df.head()

df_group = df.groupby(['Invoice Number', 'Department Long Descr', 'SUB_DEPARTMENT', 'LOCATION'])
type(df_group)
df_group.ngroups
df_group.size()
df_group.groups

# Chart of Accounts

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
    14 : [22500, 22500, 22500, 22500, 22500, 22500, 22500]}

# Create new Dataframe for the Output.
df_Output = pd.DataFrame(columns=['Entity', 'PostDate', 'DocDate', 'DocNo', 'AcctType', 'AcctNo', 'AcctName', 'Description', 'DebitAmt', 'CreditAmt', 'Loc', 'Dept', 'Provider', 'Service Line', 'Comments'])
#print(df_Output)
CoA_Index = 0
for name, r in df_group:
    #print(name[2])
    if name[2] == 'HQ':
        CoA_Index = 0
    elif name[2] == 'Lab':
        CoA_Index = 1
    elif name[2] == 'ASC':
        CoA_Index = 2
    elif name[2] == 'Clinical':
        CoA_Index = 3
    elif name[2] == 'Operating':
        CoA_Index = 4
    elif name[2] == 'NEST':
        CoA_Index = 5
    elif name[2] == 'MD':
        CoA_Index = 6
    #print(CoA_Index)
    a = name
    b = r['Gross Wages'].sum()
    c = r['OT'].sum()
    d = r['Bonus'].sum()
    e = r['Taxes - ER - Totals'].sum()
    f = r['Workers Comp Fee - Totals'].sum()
    g = r['401k/Roth-ER'].sum()
    h = r['BENEFITS wo 401K'].sum()
    i = r['TOTAL FEES'].sum()
    j = r['PTO2'].sum()
    k = r['Electronics Nontaxable'].sum()
    l = r['Reimbursement-Non Taxable'].sum()
    m = r['Total Client Charges'].sum()

    """
    ## Troubleshooting Code ##
    bigsum = b+c+d+e+f+g+h+i+j+k+l
        if abs(bigsum - m) > 1:

        print(r)
        print(a, bigsum, m)
    """

    ped = r['Pay End Date']
    ivd = r['Invoice Date']
    deptCode = r['DEPT CODE']
    dp = re.sub(r"[^0-9]","",str(deptCode)[3:10])

    
    if m != 0:
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[4][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), b, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[5][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), c, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[6][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), d, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[7][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), e, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[8][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), f, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[9][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), g, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[10][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), h, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[11][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), i, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[12][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), j, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[13][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), k, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], CoA[14][CoA_Index], "", str(name[0]) + ' ' + str(name[1]), l, "", name[3], dp, "", "", ""]
        df_Output.loc[len(df_Output.index)] = ["", str(ped)[7:17], str(ivd)[7:17], "", name[2], 23300, "", str(name[0]) + ' ' + str(name[1]), "", m, name[3], dp, "", "", ""]
    

df_Output.to_excel("zopt.xlsx", index = False)
        