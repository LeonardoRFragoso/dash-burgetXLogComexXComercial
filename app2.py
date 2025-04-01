import io
import os
import sys
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# -----------------------------------------------------------------------------
# Autenticacao e servico
# -----------------------------------------------------------------------------
gdrive_creds_path = r"C:\Users\leona\OneDrive\Documentos\dash-burgetXLogComexXComercial\gdrive_credentials.json"
scopes = ["https://www.googleapis.com/auth/drive"]

def authenticate_google_drive(credentials_path):
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    return build('drive', 'v3', credentials=creds)

drive_service = authenticate_google_drive(gdrive_creds_path)

# -----------------------------------------------------------------------------
# IDs dos arquivos existentes no Drive
# -----------------------------------------------------------------------------
budget_file_id = "1a6Z5fBd5OLg3p4ukkWpPS54slOY4HiWX"  # ID fixo da planilha BUDGET VS LOGCOMEX
logcomex_file_id = "1gyQk6l3UW-cO3HnAhTVnroFhJhe7R8VG"  # Planilha de containers gerada

# -----------------------------------------------------------------------------
# Download dos arquivos do Drive
# -----------------------------------------------------------------------------
def download_excel_file(drive_service, file_id):
    metadata = drive_service.files().get(fileId=file_id, fields='mimeType').execute()
    if metadata['mimeType'].startswith('application/vnd.google-apps'):
        request = drive_service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        print(f"Download progress: {int(status.progress() * 100)}%.")
    fh.seek(0)
    return pd.read_excel(fh)

# Ler as duas planilhas
df_budget = download_excel_file(drive_service, budget_file_id)
df_logcomex = download_excel_file(drive_service, logcomex_file_id)

# -----------------------------------------------------------------------------
# Pre-processamento
# -----------------------------------------------------------------------------
# Garantir campo de containers somados
container_cols = [
    "C20", "C40", "QTDE CONTAINER", "QTDE CONTEINER", "QUANTIDADE C20", "QUANTIDADE C40"
]
for col in container_cols:
    if col in df_logcomex.columns:
        df_logcomex[col] = pd.to_numeric(
            df_logcomex[col].astype(str).str.replace(",", "."), errors="coerce"
        ).fillna(0)

if "Containers Somados" not in df_logcomex.columns:
    df_logcomex["Containers Somados"] = df_logcomex[container_cols].sum(axis=1)

# Converter data para mes inteiro
df_logcomex['ANO/MÊS'] = pd.to_datetime(df_logcomex['ANO/MÊS'], format="%m/%Y", errors='coerce').dt.month

# Aliases de clientes para unificação
aliases = {
    # Padronização Aliança/Alianca
    "alianca": "alianca",
    "aliança": "alianca",
    # Clientes específicos
    "alianca froneri x jpa": "froneri",
    "blue water - froneri - jpa": "froneri",
    "alianca - elgin": "elgin",
    "alianca - diversos": "ambev",
    "brr reciclagem e coleta ltda": "brr reciclagem",
    "cobremax rio": "cobremax",
    "katrium industrias quimicas s.a": "katrium",
    "iff essencias e fragrancias ltda": "iff",
    "dc logistics brasil": "dc logistics",
    "ibr-lam laminacao de metais ltda": "ibr-lam",
    "alianca valgroup xerem": "valgroup",
    "alianca samsung x via varejo": "samsung",
    "alianca - cosan": "cosan",
    "alianca - ball": "ball",
    "alianca - braskem": "braskem",
    "alianca - anfrapi": "anfrapi"
}

# Palavras a serem removidas dos nomes para melhor normalização
palavras_remover = [
    "ltda", "s.a.", "s/a", "sa", "industria", "comercio",
    "x", "de", "do", "da", "dos", "das", "e"
]

def normalizar_nome(nome):
    if not isinstance(nome, str):
        return nome
    
    # Converter para minúsculo e fazer substituições básicas
    nome = (nome.lower()
               .replace("-", " ")
               .replace("_", " ")
               .replace(",", " ")
               .replace("  ", " ")
               .replace("á", "a")
               .replace("ã", "a")
               .replace("â", "a")
               .replace("ç", "c")
               .replace("é", "e")
               .replace("ê", "e")
               .replace("í", "i")
               .replace("ó", "o")
               .replace("ô", "o")
               .replace("ú", "u")
               .strip())
    
    # Remover palavras comuns
    for palavra in palavras_remover:
        nome = nome.replace(f" {palavra} ", " ")
    
    # Limpar espaços extras
    nome = " ".join(nome.split())
    return nome

def aplicar_alias(nome):
    nome_norm = normalizar_nome(nome)
    # Primeiro tenta encontrar o nome completo
    if nome_norm in aliases:
        return aliases[nome_norm]
    
    # Se não encontrar, procura por partes do nome que possam corresponder
    for key, value in aliases.items():
        if key in nome_norm:
            return value
    
    return nome_norm

# Filtrar e preparar df_logcomex
print("Processando dados do LogComex...")
df_logcomex = df_logcomex[df_logcomex['Clientes Encontrados'].notnull()].copy()

# Separar múltiplos clientes e normalizar
df_logcomex['Cliente'] = df_logcomex['Clientes Encontrados'].astype(str).str.split(',')
df_logcomex = df_logcomex.explode('Cliente')
df_logcomex['Cliente'] = df_logcomex['Cliente'].apply(lambda x: x.strip())  # Remover espaços
df_logcomex['Cliente'] = df_logcomex['Cliente'].apply(normalizar_nome)
df_logcomex['Cliente'] = df_logcomex['Cliente'].apply(aplicar_alias)

# Remover possíveis duplicatas após normalização
df_logcomex = df_logcomex.drop_duplicates(['Cliente', 'ANO/MÊS', 'Containers Somados'])

# Agrupar total de containers por cliente normalizado e mês
print("Agrupando dados do LogComex...")
df_logcomex_grouped = df_logcomex.groupby(['Cliente', 'ANO/MÊS'])['Containers Somados'].sum().reset_index()
df_logcomex_grouped['Containers Somados'] = df_logcomex_grouped['Containers Somados'].round(2)

# Aplicar normalização no budget
print("Processando dados do Budget...")
df_budget_ren = df_budget.rename(columns={"CLIENTE (BUDGET)": "Cliente", "MÊS": "ANO/MÊS"})
df_budget_ren['Cliente'] = df_budget_ren['Cliente'].apply(normalizar_nome)
df_budget_ren['Cliente'] = df_budget_ren['Cliente'].apply(aplicar_alias)

# Validação antes do merge
print("\nValidação de dados:")
clientes_budget = set(df_budget_ren['Cliente'].unique())
clientes_logcomex = set(df_logcomex_grouped['Cliente'].unique())

print("\nClientes apenas no Budget:", sorted(clientes_budget - clientes_logcomex))
print("\nClientes apenas no LogComex:", sorted(clientes_logcomex - clientes_budget))

# Remover coluna duplicada se existir
if "LogComex" in df_budget_ren.columns:
    df_budget_ren = df_budget_ren.drop(columns=["LogComex"])

# Merge alinhado por cliente e mês
print("\nRealizando merge dos dados...")
df_final = pd.merge(df_budget_ren, df_logcomex_grouped,
                    how='left', on=['Cliente', 'ANO/MÊS'])

# Renomear coluna de containers
df_final = df_final.rename(columns={"Containers Somados": "LogComex"})

# Ordenar o DataFrame final
df_final = df_final.sort_values(['Cliente', 'ANO/MÊS'])

print("\nResumo final:")
print(f"Total de clientes únicos: {len(df_final['Cliente'].unique())}")
print(f"Meses processados: {sorted(df_final['ANO/MÊS'].unique())}")

# -----------------------------------------------------------------------------
# Salvar planilha local
# -----------------------------------------------------------------------------
output_path = r"C:\Users\leona\OneDrive\Documentos\dash-burgetXLogComexXComercial\BUDGET VS LOGCOMEX (clientes).xlsx"
df_final.to_excel(output_path, index=False)

# -----------------------------------------------------------------------------
# Atualizar arquivo fixo no Google Drive com o mesmo nome e ID
# -----------------------------------------------------------------------------
fixed_output_id = "1a6Z5fBd5OLg3p4ukkWpPS54slOY4HiWX"  # Mesmo ID da planilha original

media = MediaFileUpload(output_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
drive_service.files().update(fileId=fixed_output_id, media_body=media).execute()

print("Planilha final 'BUDGET VS LOGCOMEX (clientes)' atualizada no Drive com sucesso!")