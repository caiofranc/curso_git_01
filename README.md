import pandas as pd

# Load the Excel file
file_path = r'C:\Users\SUPORTE-CAIO\Downloads\Indicadores_tecnicos.xlsx'
df = pd.read_excel(file_path)

# Group by 'técnicos' and 'categoria 3' and count the occurrences
grouped = df.groupby(['Técnico', 'Categoria 3']).size().reset_index(name='counts')

# Calculate the total counts for each 'técnicos'
total_counts = df['Técnico'].value_counts().reset_index()
total_counts.columns = ['Técnico', 'total']

# Merge the grouped data with the total counts to calculate percentages
merged = pd.merge(grouped, total_counts, on='Técnico')
merged['percentage'] = (merged['counts'] / merged['total']) * 100

# Create a report
report = merged[['Técnico', 'Categoria 3', 'counts', 'percentage']]

# Save the report to a new Excel file
report.to_excel(r'C:\Users\SUPORTE-CAIO\Downloads\Indicadores_tecnicos.xlsx', index=False)

print("O relatório foi gerado e salvo como 'relatorio_tecnicos.xlsx'.")
print(report)
