import pandas as pd

# Load the data
df = pd.read_excel('data/database.xlsx')

# Print unique values in Condición column
print("\nValores únicos en columna Condición:")
print(df['Condición'].unique())

# Print unique values in Periodo column
print("\nValores únicos en columna Periodo:")
print(df['Periodo'].unique())

# Print a sample of rows where Periodo is '1391' to check format
print("\nEjemplo de registros del periodo 1391:")
print(df[df['Periodo'] == '1391'].head())
