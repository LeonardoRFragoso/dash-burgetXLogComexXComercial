import os
import time
import win32com.client as win32
import pythoncom
from pywintypes import com_error

# Caminho do arquivo original
caminho_arquivo = r"C:\Users\leonardo.fragoso\Desktop\Projetos\dash-burgetXLogComexXComercial\iTRACKER_novo 01.06 v2.xlsx"

def aguardar_conexoes(workbook):
    print("⏳ Atualizando todas as conexões da planilha...")
    try:
        workbook.RefreshAll()
    except com_error:
        print("⚠️ Erro ao iniciar RefreshAll.")
        return

    print("⏳ Aguardando todas as conexões finalizarem...")
    while True:
        pythoncom.PumpWaitingMessages()
        atualizando = False
        for i in range(workbook.Connections.Count):
            try:
                conn = workbook.Connections.Item(i+1)
                if conn.Type == 2:  # xlConnectionTypeOLEDB = 2
                    atualizando = True
                    break
            except com_error:
                continue

        if not atualizando:
            print("✅ Todas as conexões finalizaram.")
            break

        time.sleep(2)

def abrir_e_salvar_automaticamente():
    # Inicia o Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.DisplayAlerts = False

    print("🔄 Abrindo a planilha...")
    workbook = excel.Workbooks.Open(caminho_arquivo)

    time.sleep(5)  # Pequeno delay de segurança
    aguardar_conexoes(workbook)

    print("💾 Salvando planilha original atualizada...")
    workbook.Save()

    workbook.Close(False)
    excel.Quit()
    print("✅ Processo finalizado com sucesso!")

if __name__ == "__main__":
    abrir_e_salvar_automaticamente()
