import pandas as pd

# Carregar o arquivo Excel
file_path = r'C:\Users\SUPORTE-CAIO\Downloads\Indicadores_tecnicos.xlsx'
df = pd.read_excel(file_path)

# Garantir que as colunas de data estejam no formato datetime
df['DT Fim Relato'] = pd.to_datetime(df['DT Fim Relato'], errors='coerce')
df['Criação Agend'] = pd.to_datetime(df['Criação Agend'], errors='coerce')

# Calcular a diferença em horas entre 'DT Fim Relato' e 'Criação Agend'
df['Diferença em Horas'] = (df['DT Fim Relato'] - df['Criação Agend']).dt.total_seconds() / 3600

# Calcular a média de SLA por técnico e categoria 3
# Primeiro agrupamos por Técnico e Categoria 3 e depois calculamos a média do SLA
media_sla_tecnico_categoria = df.groupby(['Técnico', 'Categoria 3'])['Diferença em Horas'].mean().reset_index()

# Renomear a coluna com a média
media_sla_tecnico_categoria.columns = ['Técnico', 'Categoria 3', 'Média SLA (Horas)']

# Adicionar o número de ocorrências por Técnico e Categoria 3
media_sla_tecnico_categoria['counts'] = df.groupby(['Técnico', 'Categoria 3']).size().values

# Calcular a média de SLA por técnico (independentemente da categoria 3)
media_sla_tecnico = df.groupby('Técnico')['Diferença em Horas'].mean().reset_index()
media_sla_tecnico.columns = ['Técnico', 'Média SLA por Técnico (Horas)']

# Mesclar as médias de SLA por Técnico e Categoria 3 com a média por técnico
report = pd.merge(media_sla_tecnico_categoria, media_sla_tecnico, on='Técnico')

# Salvar o relatório em um novo arquivo Excel
report.to_excel(r'C:\Users\SUPORTE-CAIO\Downloads\Relatorio_Medias_SLA.xlsx', index=False)

# Exibir o relatório e as médias calculadas
print("O relatório foi gerado e salvo como 'Relatorio_Medias_SLA.xlsx'.")
print(report)



______________________________________________________________________________________________________________________________________________________________________
import pandas as pd

# Carregar o arquivo Excel
file_path = r'C:\Users\SUPORTE-CAIO\Downloads\Indicadores_tecnicos.xlsx'
df = pd.read_excel(file_path)

# Garantir que as colunas de data estejam no formato datetime
df['DT Fim Relato'] = pd.to_datetime(df['DT Fim Relato'], errors='coerce')
df['Criação Agend'] = pd.to_datetime(df['Criação Agend'], errors='coerce')

# Calcular a diferença em horas entre 'DT Fim Relato' e 'Criação Agend'
df['Diferença em Horas'] = (df['DT Fim Relato'] - df['Criação Agend']).dt.total_seconds() / 3600

# Calcular a média da diferença em horas
media_diferenca_horas = df['Diferença em Horas'].mean()

# Agrupar por 'Técnico' e 'Categoria 3' e contar as ocorrências
grouped = df.groupby(['Técnico', 'Categoria 3']).size().reset_index(name='counts')

# Calcular o total de ocorrências para cada 'Técnico'
total_counts = df['Técnico'].value_counts().reset_index()
total_counts.columns = ['Técnico', 'total']

# Mesclar os dados agrupados com o total para calcular as porcentagens
merged = pd.merge(grouped, total_counts, on='Técnico')
merged['percentage'] = (merged['counts'] / merged['total']) * 100

# Criar um relatório
report = merged[['Técnico', 'Categoria 3', 'counts', 'percentage']]

# Adicionar a média da diferença em horas ao relatório (opcional, se quiser incluir)
report['Média Diferença Horas'] = media_diferenca_horas

# Salvar o relatório em um novo arquivo Excel
report.to_excel(r'C:\Users\SUPORTE-CAIO\Downloads\Indicadores_tecnicos.xlsx', index=False)

# Exibir o relatório e a média calculada
print("O relatório foi gerado e salvo como 'relatorio_tecnicos.xlsx'.")
print(report)
print(f"A média da diferença em horas é: {media_diferenca_horas:.2f}")
