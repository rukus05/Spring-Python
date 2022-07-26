# (name, $-income)
customers = [("John", 240000),
             ("Alice", 120000),
             ("Ann", 1100000),
             ("Zach", 44000)]

# your high-value customers earning <$1M
whales = []
whales = [x for x,y in customers if y>1000000]
print(whales)
print(customers[len(customers) - 2])

# ['Ann']

mos = input("How many months are you running this report for?")
#moslist = [None] * int(mos)
moslist = list(range(int(mos)))



