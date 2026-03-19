import win32com.client
import time
import pyperclip

# =========================================
# CONEXÃO COM SAP
# =========================================
def conectar_sap():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    return session


# =========================================
# 📊 KSB1
# =========================================
def executar_ksb1(session):

    session.findById("wnd[0]/tbar[0]/okcd").text = "KSB1"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    # 🔴 AJUSTAR ID DO CAMPO ABI1
    # session.findById("wnd[0]/usr/ctxtXXXX").text = "ABI1"

    # Variante (Shift + F5)
    session.findById("wnd[0]").sendVKey(17)
    time.sleep(2)

    # Criado por
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "99846147"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Executar
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(5)


# =========================================
# 📤 EXPORTAR SAP → EXCEL
# =========================================
def exportar_excel(session):

    # 🔴 AJUSTAR CAMINHO DO MENU (usar gravação SAP)
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    time.sleep(2)

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(2)

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(3)


# =========================================
# 📊 TRATAR EXCEL 1 (KSB1)
# =========================================
def tratar_excel_ksb1():

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.ActiveWorkbook
    ws = wb.ActiveSheet

    # Converter para número
    ws.UsedRange.Value = ws.UsedRange.Value

    # Filtrar coluna C (Ordem)
    ws.Range("A1").AutoFilter(Field=3, Criteria1="=")

    # Deletar vazios
    try:
        ws.Range("A2:A100000").SpecialCells(12).EntireRow.Delete()
    except:
        pass

    ws.AutoFilterMode = False

    # Copiar ordens
    last_row = ws.Cells(ws.Rows.Count, 3).End(-4162).Row

    ordens = []
    for i in range(2, last_row + 1):
        valor = ws.Cells(i, 3).Value
        if valor:
            ordens.append(str(valor))

    pyperclip.copy("\n".join(ordens))


# =========================================
# 📊 KOB1
# =========================================
def executar_kob1(session):

    session.findById("wnd[0]/tbar[0]/okcd").text = "KOB1"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    # Variante
    session.findById("wnd[0]").sendVKey(17)
    time.sleep(2)

    session.findById("wnd[1]/usr/txtENAME-LOW").text = "99846147"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Seleções múltiplas
    # 🔴 AJUSTAR ID DO BOTÃO ORDEM
    session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press()
    time.sleep(2)

    # Upload clipboard (Shift + F12)
    session.findById("wnd[1]").sendVKey(24)

    # Transferir
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Executar
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(5)


# =========================================
# 📊 TRATAR EXCEL 2 (KOB1)
# =========================================
def tratar_excel_kob1():

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.ActiveWorkbook
    ws = wb.ActiveSheet

    # Filtrar coluna D (Centro de Custo vazio)
    ws.Range("A1").AutoFilter(Field=4, Criteria1="=")

    try:
        ws.Range("A2:A100000").SpecialCells(12).EntireRow.Delete()
    except:
        pass

    ws.AutoFilterMode = False

    # Ordenar decrescente
    ws.Range("A1").Sort(Key1=ws.Range("D2"), Order1=2)

    # Preencher vazios com valor anterior
    last_row = ws.Cells(ws.Rows.Count, 4).End(-4162).Row

    for i in range(2, last_row + 1):
        if not ws.Cells(i, 4).Value:
            ws.Cells(i, 4).Value = ws.Cells(i - 1, 4).Value


# =========================================
# 🚀 FLUXO PRINCIPAL
# =========================================
def main():

    session = conectar_sap()

    print("Executando KSB1...")
    executar_ksb1(session)

    print("Exportando KSB1...")
    exportar_excel(session)

    print("Tratando Excel KSB1...")
    tratar_excel_ksb1()

    print("Executando KOB1...")
    executar_kob1(session)

    print("Exportando KOB1...")
    exportar_excel(session)

    print("Tratando Excel KOB1...")
    tratar_excel_kob1()

    # =====================================
    # 🔴 CONTINUAÇÃO DO SEU ROBÔ AQUI
    # =====================================
    # Exemplo:
    # chamar nova transação
    # tratar mais planilhas
    # salvar arquivos
    # enviar email
    # etc


if __name__ == "__main__":
    main()
