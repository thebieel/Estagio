import pandas as pd
import matplotlib.pyplot as plt

# Caminho do arquivo Excel (ajuste conforme o local em que ele está salvo)
arquivo = 'C:\\Users\\gabriel.pontes\\Documents\\Ligações x Agendamentos.xlsx'

# Ler a planilha, selecionando a aba correta
df = pd.read_excel(arquivo, sheet_name='Relatório')

# Selecionar as colunas desejadas (Data, Agendamentos e Ligações)
df_plot = df.iloc[:, [0, 4, 5]].copy()
df_plot.columns = ['Data', 'Agendamentos', 'Ligações']
df_plot['Data'] = pd.to_datetime(df_plot['Data'])

# Criar o gráfico
fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(df_plot['Data'], df_plot['Agendamentos'], marker='o', label='Agendamentos')
ax.plot(df_plot['Data'], df_plot['Ligações'], marker='o', label='Ligações Recebidas')

# Adicionar rótulos de dados
for i, row in df_plot.iterrows():
    ax.text(row['Data'], row['Agendamentos'] + 3, str(row['Agendamentos']), ha='center', fontsize=8)
    ax.text(row['Data'], row['Ligações'] + 3, str(row['Ligações']), ha='center', fontsize=8)

# Ajustar o layout do gráfico
ax.set_title('Agendamentos vs Ligações Recebidas por Dia')
ax.set_xlabel('Data')
ax.set_ylabel('Quantidade')
ax.legend()
ax.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()

# Mostrar o gráfico
plt.show()
