import io
import os
import sys
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# -----------------------------------------------------------------------------
# Configurações de credenciais e autenticação no Google Drive
# -----------------------------------------------------------------------------
gdrive_creds_path = r"C:\Users\leona\OneDrive\Documentos\dash-burgetXLogComexXComercial\gdrive_credentials.json"
scopes = ["https://www.googleapis.com/auth/drive"]

def authenticate_google_drive(credentials_path):
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    drive_service = build('drive', 'v3', credentials=creds)
    return drive_service

drive_service_upload = authenticate_google_drive(gdrive_creds_path)
drive_service_dummy = authenticate_google_drive(gdrive_creds_path)

# -----------------------------------------------------------------------------
# IDs dos arquivos para download
# -----------------------------------------------------------------------------
file_id2 = "1OFGkiN66JVZELUbSIod_yctInQceDltn"

def download_excel_file(drive_service, file_id):
    file_metadata = drive_service.files().get(fileId=file_id, fields='mimeType').execute()
    mime_type = file_metadata.get('mimeType', '')
    if mime_type.startswith('application/vnd.google-apps'):
        request = drive_service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        print(f"Download progress: {int(status.progress() * 100)}%.")
    fh.seek(0)
    return fh

# -----------------------------------------------------------------------------
# Download e leitura do arquivo Excel
# -----------------------------------------------------------------------------
excel_file2 = download_excel_file(drive_service_dummy, file_id2)
df_full = pd.read_excel(excel_file2)

# -----------------------------------------------------------------------------
# Pré-processamento
# -----------------------------------------------------------------------------
df_full = df_full[df_full['Clientes Encontrados'].notnull() & (df_full['Clientes Encontrados'] != '')].copy()

# Localizar a coluna de "ANO/MÊS" de forma flexível
coluna_ano_mes = None
for col in df_full.columns:
    if col.strip().lower().replace("ê", "e") in ["ano/mes", "ano/mes", "ano_mes"]:
        coluna_ano_mes = col
        break

if not coluna_ano_mes:
    print("Coluna 'ANO/MÊS' não encontrada de forma flexível na planilha.")
    sys.exit(1)
else:
    df_full.rename(columns={coluna_ano_mes: "ANO/MÊS"}, inplace=True)

# Conversão para campos de containers
for col in ["C20", "C40", "QUANTIDADE C20", "QUANTIDADE C40"]:
    if col in df_full.columns:
        df_full[col] = df_full[col].astype(str).str.replace(',', '.').str.strip()
        df_full[col] = pd.to_numeric(df_full[col], errors='coerce').fillna(0)

# Agrupar diretamente por Categoria, Cliente e ANO/MÊS, somando todos os campos disponíveis
agrupado = df_full.groupby(["Categoria", "Clientes Encontrados", "ANO/MÊS"], as_index=False)[
    [col for col in ["C20", "C40", "QUANTIDADE C20", "QUANTIDADE C40"] if col in df_full.columns]
].sum()

# Calcular Containers Somados
agrupado["Containers Somados"] = agrupado[
    [col for col in ["C20", "C40", "QUANTIDADE C20", "QUANTIDADE C40"] if col in agrupado.columns]
].sum(axis=1)

# Formatar ANO/MÊS para MM/YYYY
agrupado["ANO/MÊS"] = pd.to_datetime(agrupado["ANO/MÊS"], format="%Y%m", errors='coerce').dt.strftime("%m/%Y")

# Reordenar colunas
colunas_finais = ["Categoria", "Clientes Encontrados", "ANO/MÊS"] + \
                 [col for col in ["C20", "C40", "QUANTIDADE C20", "QUANTIDADE C40"] if col in agrupado.columns] + \
                 ["Containers Somados"]

final_df = agrupado[colunas_finais]

# -----------------------------------------------------------------------------
# Salvar planilha final localmente
# -----------------------------------------------------------------------------
output_path = r"C:\Users\leona\OneDrive\Documentos\dash-burgetXLogComexXComercial\ContainersClientesMes.xlsx"
final_df.to_excel(output_path, index=False)
print("Planilha final gerada com sucesso!")

# -----------------------------------------------------------------------------
# Atualizar no Google Drive
# -----------------------------------------------------------------------------
fixed_file_id = "1gyQk6l3UW-cO3HnAhTVnroFhJhe7R8VG"

def update_file_on_drive(drive_service, file_path, file_id, mime_type):
    try:
        drive_service.files().get(fileId=file_id).execute()
    except Exception as e:
        print(f"Erro: Arquivo com ID {file_id} não encontrado.")
        sys.exit(1)
    media = MediaFileUpload(file_path, mimetype=mime_type)
    updated_file = drive_service.files().update(fileId=file_id, media_body=media).execute()
    return updated_file

update_file_on_drive(
    drive_service_upload,
    output_path,
    fixed_file_id,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

print("Arquivo atualizado no Google Drive com sucesso!")
