import pandas as pd

with open('lc.txt', 'r') as file:
    lines = file.readlines()

# Split each line into LC code and description using tab as the delimiter
data = [line.strip().split('\t', 1) for line in lines]

df = pd.DataFrame(data, columns=['Code', 'Description'])
df['Code'] = df['Code'].str.strip()
df['Description'] = df['Description'].str.strip()

print(df)
df.to_excel('lc-codes.xlsx', index=False)