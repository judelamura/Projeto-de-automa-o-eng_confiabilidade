# Importar as bibliotecas necessárias
import openpyxl
import pandas as pd
from sqlalchemy import create_engine
from time import sleep
import pyautogui

# Conectar ao banco de dados SAP HANA
hana_engine = create_engine('hana://username:password@host:port/database')

# Acessar a transação KSB1 e inserir o código
codigo = "SEU_CODIGO"
query = f"CALL KSB1('{codigo}', '')"
with hana_engine.connect() as conn:
    conn.execute(query)
    sleep(5)  # Aguardar a execução da transação

# Chamar a variante usando o atalho Shift+F5
pyautogui.hotkey('shift', 'f5')
sleep(5)  # Aguardar a execução da variante

# Ler os dados do SAP HANA
query = "SELECT * FROM sua_tabela_no_hana"
df = pd.read_sql_query(query, hana_engine)

# Carregar a planilha existente
workbook = openpyxl.load_workbook('sua_planilha.xlsx')
sheet = workbook['Nome_da_aba']

# Atualizar a planilha com os dados do SAP HANA
for i, row in df.iterrows():
    for j, col in enumerate(df.columns):
        sheet.cell(row=i+2, column=j+1, value=row[col])

# Salvar a planilha atualizada
workbook.save('sua_planilha_atualizada.xlsx')
print("Planilha atualizada com sucesso!")
