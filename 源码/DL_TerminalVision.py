import random
import pandas as pd

df = pd.read_excel(r'./Sources/list.xlsx', sheet_name='list')
data = format(df.values)
stu_num = len(df)
rand = random.randint(0, stu_num-1)
print(df['list'][rand])