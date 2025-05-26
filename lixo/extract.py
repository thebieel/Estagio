import pandas as pd

# Caminho do arquivo Excel (ajuste para o seu)
arquivo = 'C:\\Users\\gabriel.pontes\\Documents\\Ligações x Agendamentos.xlsx'

# Ler a planilha 'Relatório'
df = pd.read_excel(arquivo, sheet_name='Relatório')

# Selecionar colunas e renomear
df_plot = df.iloc[:, [0, 4, 5]].copy()
df_plot.columns = ['Data', 'Agendamentos', 'Ligações']
df_plot['Data'] = pd.to_datetime(df_plot['Data'])

# Gerar array de datas em formato string (YYYY-MM-DD)
labels = df_plot['Data'].dt.strftime('%Y-%m-%d').tolist()

# Gerar arrays de valores para agendamentos e ligações
agendamentos = df_plot['Agendamentos'].tolist()
ligacoes = df_plot['Ligações'].tolist()

# Imprimir arrays em formato JS
print("labels = ", labels)
print("agendamentos = ", agendamentos)
print("ligacoes = ", ligacoes)
