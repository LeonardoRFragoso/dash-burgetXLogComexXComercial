import os
import time
import io
import pythoncom
import win32com.client as win32
from pywintypes import com_error
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# === CONFIGURAÇÕES ===
FILE_ID = "1wV5NQQYLhYOq0jdCCGTktjsNTLRZuZRE"  # ID da planilha no GDrive
NOME_ARQUIVO = "iTRACKER_novo 01.06 v2.xlsx"
CAMINHO_LOCAL = os.path.join(os.getcwd(), NOME_ARQUIVO)

# === AUTENTICAÇÃO GDRIVE ===
def autenticar_drive():
    SCOPES = ['https://www.googleapis.com/auth/drive']
    SERVICE_ACCOUNT_FILE = 'service_account.json'

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('drive', 'v3', credentials=credentials)
    return service

# === DOWNLOAD DO GDRIVE ===
def baixar_planilha(service):
    request = service.files().get_media(fileId=FILE_ID)
    fh = io.FileIO(CAMINHO_LOCAL, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    print("⬇️ Baixando planilha do Google Drive...")
    while not done:
        status, done = downloader.next_chunk()
        print(f"Progresso: {int(status.progress() * 100)}%")
    print("✅ Download concluído.")

# === ATUALIZAÇÃO NO EXCEL ===
def aguardar_conexoes(workbook):
    print("⏳ Aguardando conexões finalizarem...")
    while True:
        pythoncom.PumpWaitingMessages()
        atualizando = False
        for i in range(workbook.Connections.Count):
            try:
                conn = workbook.Connections.Item(i+1)
                if conn.Type == 2:
                    atualizando = True
                    break
            except com_error:
                continue
        if not atualizando:
            print("✅ Conexões finalizadas.")
            break
        time.sleep(2)


def atualizar_excel():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.DisplayAlerts = False

    print("🔄 Abrindo planilha no Excel...")
    workbook = excel.Workbooks.Open(CAMINHO_LOCAL)
    time.sleep(5)

    # Focar na janela do Excel antes de enviar teclas
    workbook.Activate()
    excel.Windows(workbook.Name).Activate()
    time.sleep(1)

    # Disparar atualização via ALT+F5
    print("⌨️ Enviando ALT+F5 para atualizar dados no Excel...")
    excel.SendKeys("%{F5}", True)

    # Aguardar conexões (Power Query, etc.) finalizarem
    aguardar_conexoes(workbook)

    print("💾 Salvando alterações...")
    workbook.Save()
    workbook.Close(False)
    excel.Quit()
    print("✅ Planilha atualizada com sucesso.")

# === UPLOAD PARA O GDRIVE ===
def enviar_para_drive(service):
    media = MediaFileUpload(
        CAMINHO_LOCAL,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )
    print("📤 Enviando planilha atualizada ao Google Drive...")
    service.files().update(fileId=FILE_ID, media_body=media).execute()
    print("✅ Upload concluído.")

# === EXECUÇÃO COMPLETA ===
def main():
    service = autenticar_drive()
    baixar_planilha(service)
    atualizar_excel()
    enviar_para_drive(service)

if __name__ == "__main__":
    main()
