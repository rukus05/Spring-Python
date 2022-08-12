import pandas as pd
df = pd.read_excel("002-allJuly-alloc.xlsx")
#df.head()

df_group = df.groupby(['Invoice Number', 'Department Long Descr', 'SUB_DEPARTMENT'])
type(df_group)
df_group.ngroups
df_group.size()
df_group.groups
for name, r in df_group:
    if name[2] == 'HQ':
        a = name
        b = r['Gross Wages'].sum()
        c = r['Employee Name']
        print(a, b)
        #print(r['Employee Name'])
    
#    if inv[2] == 'HQ':
#        print(inv)
#z = df_group['Gross Wages'].sum()
#z
#for inv in z:
#    print(inv)
#if z['Invoice Number'] == 5909273:
#    print(z['Department Long Descr'])