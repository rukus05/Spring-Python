import pandas as pd

data = [[10, 18, 11], [13, 15, 8], [9, 20, 3]]

df = pd.DataFrame(data)
print(df)
dict_test = {}
dict_test = df.sum()
print(dict_test[1])