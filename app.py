import io
import os
import pandas as pd
import re
import logging
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# Configuração do logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler()]
)

#--------------------------------------
# CONFIGURAÇÕES E CONSTANTES (Script 1)
BUDGET_FILE_ID = "1oStsqOKHcLd-xl7bRsS5ewWhJ4P3mWW9bo4Ux86FTyQ"
LOGCOMEX_FILE_ID = "1OFGkiN66JVZELUbSIod_yctInQceDltn"
FINAL_COMPARATIVO_FILE_ID = "13a6L40GmbG_bmPGKMx6W_WYO9iW8xW0J"
CRED_PATH = "service_account.json"
CLIENTES_TXT = "clientes.txt"  # arquivo de clientes (utilizado no Script 1)

#--------------------------------------
# CONFIGURAÇÕES (Script 2)
ITRACKER_PATH = r"C:\Users\leonardo.fragoso\Desktop\Projetos\dash-burgetXLogComexXComercial\iTRACKER_novo 01.06 v2.xlsx"
CLIENTES_TXT_SCRIPT2 = r"C:\Users\leonardo.fragoso\Desktop\Projetos\dash-burgetXLogComexXComercial\clientes.txt"
COMPARATIVO_PATH = r"C:\Users\leonardo.fragoso\Desktop\Projetos\dash-burgetXLogComexXComercial\comparativo_budget_vs_logcomex_final.xlsx"

#--------------------------------------
# NOVA CONFIGURAÇÃO: IDs de destino para os envios ao Google Drive
FINAL_ATUALIZADO_FILE_ID = "1Bphi7lChPqh12kAStpupXJmCbwcdImKo"       # Para comparativo_final_atualizado.xlsx
INTERMEDIARY_FILE_ID    = "1uYCQ9wrWwTqocnF3qxvX7uFwrjqv93cf"       # Para contagem_por_cliente.xlsx

#--------------------------------------
# FUNÇÕES COMUNS (Script 1 e Script 2)
def normalizar_nome(nome):
    if not isinstance(nome, str):
        return nome
    nome = nome.lower().replace("-", " ").replace("_", " ").replace(",", " ")
    nome = re.sub(r'\b\d{6,}\b', '', nome)
    for a, b in {"á": "a", "ã": "a", "â": "a", "ç": "c", "é": "e", "ê": "e", "í": "i", "ó": "o", "ô": "o", "ú": "u"}.items():
        nome = nome.replace(a, b)
    remover = [
        "ltda", "s.a.", "s/a", "sa", "industria", "comercio", "x", "de",
        "do", "da", "dos", "das", "e", "importacao", "exportacao",
        "importadora", "distribuidora", "em recuperacao judicial", "do brasil", "abr"
    ]
    for palavra in remover:
        nome = re.sub(rf'\b{re.escape(palavra)}\b', '', nome)
    return " ".join(nome.split())

def aplicar_keywords_match(nome, keywords):
    """
    Atualizada para considerar múltiplas correspondências parciais:
    Verifica se cada token da keyword está presente no nome normalizado.
    Retorna a keyword com maior comprimento (mais completa).
    """
    nome_norm = normalizar_nome(nome)
    matches = [kw for kw in keywords if all(token in nome_norm.split() for token in kw.split())]
    if matches:
        return max(matches, key=len)
    return nome_norm

#--------------------------------------
# FUNÇÕES ESPECÍFICAS DO SCRIPT 1 (Google Drive e Processamento Budget x LogComex)
def authenticate_drive():
    logging.info("Autenticando no Google Drive...")
    creds = Credentials.from_service_account_file(CRED_PATH, scopes=["https://www.googleapis.com/auth/drive"])
    return build("drive", "v3", credentials=creds)

def download_excel(service, file_id):
    logging.info("Iniciando download do arquivo do Drive (ID: %s)...", file_id)
    metadata = service.files().get(fileId=file_id, fields="mimeType").execute()
    mime = metadata.get("mimeType", "")
    if mime == "application/vnd.google-apps.spreadsheet":
        request = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        logging.debug("Download %d%% concluído.", int(status.progress() * 100))
    fh.seek(0)
    logging.info("Download do arquivo (ID: %s) concluído.", file_id)
    return pd.read_excel(fh)

def extrair_mes(row):
    if pd.notna(row.get("ANO/MÊS")):
        val = str(row["ANO/MÊS"])
        if len(val) == 6:
            return int(val[-2:])
    for col in ["DATA DE EMBARQUE", "DATA EMBARQUE", "ETA", "ETS"]:
        if col in row and pd.notna(row[col]):
            dt = pd.to_datetime(row[col], errors="coerce")
            if pd.notna(dt):
                return dt.month
    return None

def update_file_in_drive(service, file_id, local_file_path, mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    if not os.path.exists(local_file_path):
        logging.error("Arquivo não encontrado: %s", local_file_path)
        return False
    file_size = os.path.getsize(local_file_path)
    if file_size == 0:
        logging.error("Arquivo está vazio: %s", local_file_path)
        return False
    media = MediaFileUpload(local_file_path, mimetype=mime_type, resumable=True)
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            service.files().update(fileId=file_id, media_body=media).execute()
            logging.info("Arquivo atualizado com sucesso: %s", local_file_path)
            return True
        except Exception as e:
            retry_count += 1
            logging.warning("Tentativa %d falhou para %s. Erro: %s", retry_count, local_file_path, str(e))
    return False

def extrair_e_normalizar_nomes_logcomex(df_log, colunas_clientes, keywords):
    """
    Extrai os nomes das colunas especificadas da planilha LogComex, normaliza-os
    e retorna um dicionário com a correspondência obtida pela função aplicar_keywords_match.
    """
    nomes = set()
    for col in colunas_clientes:
        if col in df_log.columns:
            df_log[col].dropna().apply(lambda x: [nomes.add(normalizar_nome(nome.strip())) for nome in str(x).split(",") if nome.strip()])
    correspondencias = {nome: aplicar_keywords_match(nome, keywords) for nome in nomes}
    return correspondencias

#--------------------------------------
# Mapeamento de nomes comerciais internos e outros agrupamentos manuais
nomes_comerciais = {
    "rio janeiro refrescos": "coca cola andina",
    "pif paf": "rio branco alimentos",
    "iff": "iff essencias fragrancias",
    "iff taubate": "iff essencias fragrancias",
    "iff guadalupe": "iff essencias fragrancias",
    "katrium honorio": "katrium industrias quimicas",
    "katrium santa cruz": "katrium industrias quimicas",
    "katrium": "katrium industrias quimicas",
    "maersk": "alianca naval empresa navegacao",  # Novo agrupamento: maersk passa a ser parte de alianca

    # Novos agrupamentos conforme solicitado:
    "blue water logistics": "blue water",
    "blue water shipping brasil": "blue water",
    "art bag rio": "art bag",
    "katrium industrias quimicas s.a.": "katrium industrias quimicas",
    "nov wellbore technologies brasil equipamentos servicos": "nov flexibles equipamentos servicos .",
    "seb brasil produtos domesticos": "seb brasil prods.dom.."
}

#--------------------------------------
# FUNÇÃO PRINCIPAL DO SCRIPT 1 (Processamento Budget x LogComex)
def main_script1():
    """
    Processa os dados do Budget e LogComex, agrupando os registros,
    e gera a planilha "comparativo_budget_vs_logcomex_final.xlsx" contendo todas
    as empresas do Budget.
    """
    logging.info("Iniciando Script 1: Processamento Budget vs LogComex")
    service = authenticate_drive()
    logging.info("Fazendo download dos arquivos do Google Drive...")
    df_budget = download_excel(service, BUDGET_FILE_ID)
    df_log = download_excel(service, LOGCOMEX_FILE_ID)
    
    with open(CLIENTES_TXT, "r", encoding="utf-8") as f:
        keywords = [normalizar_nome(l.strip()) for l in f if l.strip()]
    
    logging.info("Processando dados do Budget...")
    df_budget["cliente_can"] = df_budget["CLIENTE (BUDGET)"].apply(lambda x: aplicar_keywords_match(x, keywords))
    # Agrupa mantendo todas as empresas do Budget
    df_budget_grouped = df_budget.groupby(["cliente_can", "MÊS"]).agg({"BUDGET": "sum"}).reset_index()
    
    logging.info("Processando dados do LogComex...")
    container_cols = ["C20", "C40", "QTDE CONTAINER", "QTDE CONTEINER", "QUANTIDADE C20", "QUANTIDADE C40"]
    for col in container_cols:
        if col in df_log.columns:
            df_log[col] = pd.to_numeric(df_log[col].astype(str).str.replace(",", "."), errors="coerce").fillna(0)
    df_log["Containers Somados"] = df_log[[col for col in container_cols if col in df_log.columns]].sum(axis=1)
    df_log["MÊS"] = df_log.apply(extrair_mes, axis=1)
    
    colunas_clientes = [
        "AGENTE DE CARGA", "AGENTE INTERNACIONAL", "ARMADOR",
        "CONSIGNATARIO FINAL", "CONSIGNATÁRIO", "CONSOLIDADOR",
        "DESTINATÁRIO", "NOME EXPORTADOR", "NOME IMPORTADOR",
        "REMETENTE", "Clientes Encontrados"
    ]
    
    # Extrair nomes do LogComex e normalizar
    correspondencias_logcomex = extrair_e_normalizar_nomes_logcomex(df_log, colunas_clientes, keywords)
    registros = []
    for _, row in df_log.iterrows():
        nomes = set()
        for col in colunas_clientes:
            if col in df_log.columns and pd.notna(row[col]):
                nomes.update([p.strip() for p in str(row[col]).split(",") if p.strip()])
        for nome in nomes:
            registros.append({
                "cliente_log": nome,
                "cliente_can": correspondencias_logcomex.get(normalizar_nome(nome), aplicar_keywords_match(nome, keywords)),
                "MÊS": row["MÊS"],
                "Containers Somados": row.get("Containers Somados", 0),
                "Categoria": row.get("Categoria", "")
            })
    df_log_grouped = pd.DataFrame(registros).groupby(["cliente_can", "MÊS", "Categoria"])["Containers Somados"].sum().reset_index()
    
    # Filtra para manter somente os clientes presentes no Budget
    clientes_budget = set(df_budget_grouped["cliente_can"]) & set(keywords)
    df_budget_final = df_budget_grouped[df_budget_grouped["cliente_can"].isin(clientes_budget)]
    df_log_final = df_log_grouped[df_log_grouped["cliente_can"].isin(clientes_budget)]
    
    logging.info("Realizando pivot dos dados do LogComex...")
    df_log_pivot = df_log_final.pivot_table(
        index=["cliente_can", "MÊS"],
        columns="Categoria",
        values="Containers Somados",
        fill_value=0
    ).reset_index()
    for categoria in ["Importação", "Exportação", "Cabotagem"]:
        if categoria not in df_log_pivot.columns:
            df_log_pivot[categoria] = 0

    # Merge Budget e LogComex utilizando join à esquerda para manter todas as empresas do Budget
    df_final = pd.merge(df_budget_final, df_log_pivot, how="left", on=["cliente_can", "MÊS"]).fillna(0)
    df_final["Cliente"] = df_final["cliente_can"]
    df_final["BUDGET"] = df_final["BUDGET"].astype(int)
    for categoria in ["Importação", "Exportação", "Cabotagem"]:
        df_final[categoria] = df_final[categoria].round().astype(int)
    df_final = df_final[["Cliente", "MÊS", "BUDGET", "Importação", "Exportação", "Cabotagem"]].sort_values(by=["Cliente", "MÊS"])
    
    # Salva a planilha comparativa final de Budget vs LogComex
    final_filename = "comparativo_budget_vs_logcomex_final.xlsx"
    df_final.to_excel(final_filename, index=False)
    logging.info("Arquivo gerado: %s", final_filename)
    logging.info("Atualizando a planilha final do Budget vs LogComex no Google Drive...")
    drive_service = authenticate_drive()
    sucesso_upload = update_file_in_drive(drive_service, FINAL_COMPARATIVO_FILE_ID, final_filename)
    if sucesso_upload:
        logging.info("Planilha final do Budget vs LogComex atualizada com sucesso no Google Drive.")
    else:
        logging.error("Falha ao atualizar a planilha final do Budget vs LogComex no Google Drive.")

#--------------------------------------
# FUNÇÃO PRINCIPAL DO SCRIPT 2 (Processamento iTRACKER e Merge Completo)
def main_script2():
    """
    Processa a planilha iTRACKER (aba "Planilha1") e faz um merge full (outer)
    com a planilha comparativa final obtida do Script 1, garantindo que todas as
    empresas de todas as fontes sejam incluídas.
    Considera somente registros de 2025 a partir do mês 4.
    """
    logging.info("Iniciando processamento do iTRACKER e merge completo com a base comparativa.")
    path_itracker = ITRACKER_PATH
    path_clientes_txt = CLIENTES_TXT_SCRIPT2
    path_comparativo = COMPARATIVO_PATH

    with open(path_clientes_txt, "r", encoding="utf-8") as f:
        keywords = [normalizar_nome(l.strip()) for l in f if l.strip()]
    
    logging.info("Lendo dados da planilha iTRACKER (Planilha1)...")
    df_itracker = pd.read_excel(path_itracker, sheet_name="Planilha1")
    df_itracker["DataEmissao"] = pd.to_datetime(df_itracker["DataEmissao"], errors='coerce')
    
    logging.info("Aplicando filtros obrigatórios na planilha iTRACKER...")
    cond_empresa = df_itracker["Empresa"] == "IRB MATRIZ"
    cond_tipo_atendimento = df_itracker["TipoAtendimento"] == "ATENDIMENTO"
    cond_tipo_nota = (
        (df_itracker["tiponotafiscal"] == "Nota Fiscal") |
        (df_itracker["tiponotafiscal"].isna()) |
        (df_itracker["tiponotafiscal"].astype(str).str.strip() == "")
    )
    cond_status = df_itracker["Status"] == "autorizado"
    df_filtrado = df_itracker[cond_empresa & cond_tipo_atendimento & cond_tipo_nota & cond_status].copy()
    logging.info("Número de registros após filtros: %d", df_filtrado.shape[0])
    
    # Aplica a correspondência para agrupar os clientes do iTRACKER
    df_filtrado["cliente_can"] = df_filtrado["cliente"].apply(lambda x: aplicar_keywords_match(x, keywords))
    df_filtrado["MÊS"] = df_filtrado["DataEmissao"].dt.month
    df_filtrado["ANO"] = df_filtrado["DataEmissao"].dt.year
    # Considera somente o ano de 2025 a partir do mês 4
    df_filtrado = df_filtrado[(df_filtrado["ANO"] == 2025) & (df_filtrado["MÊS"] >= 4)]
    df_contagem = df_filtrado.groupby(["cliente_can", "ANO", "MÊS"]).size().reset_index(name="Quantidade")
    df_contagem = df_contagem.rename(columns={"cliente_can": "Cliente"})
    
    # Salva a contagem para apoio
    df_contagem.to_excel("contagem_por_cliente.xlsx", index=False)
    logging.info("Arquivo de apoio 'contagem_por_cliente.xlsx' gerado com sucesso!")
    logging.info("Enviando arquivo de apoio 'contagem_por_cliente.xlsx' para o Google Drive...")
    drive_service = authenticate_drive()
    sucesso_upload_intermediary = update_file_in_drive(drive_service, INTERMEDIARY_FILE_ID, "contagem_por_cliente.xlsx")
    if sucesso_upload_intermediary:
        logging.info("Arquivo 'contagem_por_cliente.xlsx' atualizado com sucesso no Google Drive.")
    else:
        logging.error("Falha ao atualizar o arquivo 'contagem_por_cliente.xlsx' no Google Drive.")
    
    logging.info("Lendo a planilha comparativa final gerada (Budget vs LogComex)...")
    df_comparativo = pd.read_excel(path_comparativo)
    if "ANO" not in df_comparativo.columns:
        logging.warning("Planilha comparativa não tem coluna ANO. Adicionando ANO=2025 como padrão.")
        df_comparativo["ANO"] = 2025
    
    # Merge outer: inclui todas as empresas de todas as fontes
    df_merged = pd.merge(df_comparativo, df_contagem, on=["Cliente", "ANO", "MÊS"], how="outer")
    df_merged.fillna(0, inplace=True)
    df_merged["BUDGET"] = df_merged["BUDGET"].astype(int)
    if "Importação" in df_merged.columns:
        df_merged["Importação"] = df_merged["Importação"].astype(int)
    if "Exportação" in df_merged.columns:
        df_merged["Exportação"] = df_merged["Exportação"].astype(int)
    if "Cabotagem" in df_merged.columns:
        df_merged["Cabotagem"] = df_merged["Cabotagem"].astype(int)
    df_merged["Quantidade_iTRACKER"] = df_merged["Quantidade"].astype(int)
    df_merged.drop(columns=["Quantidade"], inplace=True)
    
    # Remapeia os nomes com os mapeamentos comerciais (incluindo maersk e os novos agrupamentos)
    df_merged["Cliente"] = df_merged["Cliente"].replace(nomes_comerciais)
    
    # Reagrupa para consolidar possíveis duplicações (mesmo Cliente e MÊS)
    df_merged = df_merged.groupby(["Cliente", "MÊS"]).agg({
        "BUDGET": "sum",
        "Importação": "sum",
        "Exportação": "sum",
        "Cabotagem": "sum",
        "Quantidade_iTRACKER": "sum"
    }).reset_index()
    
    # Adição das novas colunas de métricas:
    # Total de oportunidades = Importação + Exportação + Cabotagem
    df_merged["Total Oportunidades"] = df_merged["Importação"] + df_merged["Exportação"] + df_merged["Cabotagem"]
    
    # Aproveitamento de Oportunidade (%) = (Quantidade_iTRACKER / Total Oportunidades) * 100
    df_merged["Aproveitamento de Oportunidade (%)"] = (
        (df_merged["Quantidade_iTRACKER"] / df_merged["Total Oportunidades"].replace(0, 1)) * 100
    ).round(2)
    df_merged.loc[df_merged["Total Oportunidades"] == 0, "Aproveitamento de Oportunidade (%)"] = 0
    
    # Realização do Budget (%) = (Quantidade_iTRACKER / BUDGET) * 100
    df_merged["Realização do Budget (%)"] = (
        (df_merged["Quantidade_iTRACKER"] / df_merged["BUDGET"].replace(0, 1)) * 100
    ).round(2)
    df_merged.loc[df_merged["BUDGET"] == 0, "Realização do Budget (%)"] = 0
    
    # Desvio Budget vs Oportunidade (%) = ((Total Oportunidades - BUDGET) / BUDGET) * 100
    df_merged["Desvio Budget vs Oportunidade (%)"] = (
        ((df_merged["Total Oportunidades"] - df_merged["BUDGET"]) / df_merged["BUDGET"].replace(0, 1)) * 100
    ).round(2)
    df_merged.loc[df_merged["BUDGET"] == 0, "Desvio Budget vs Oportunidade (%)"] = 0

    # Novas colunas de acompanhamento diário (considerando mês com 30 dias)
    # Target Diário Esperado = BUDGET / 30  (renomeado de "Budget Diário Esperado")
    df_merged["Target Diário Esperado"] = (df_merged["BUDGET"] / 30).round(2)
    # Target Acumulado = Target Diário Esperado * Dia Atual (meta acumulada até hoje)
    current_day = pd.Timestamp.now().day
    df_merged["Target Acumulado"] = (df_merged["Target Diário Esperado"] * current_day).round(2)
    # Gap de Realização = Target Acumulado - Quantidade_iTRACKER
    df_merged["Gap de Realização"] = (df_merged["Target Acumulado"] - df_merged["Quantidade_iTRACKER"]).round(2)
    
    # Organiza as colunas e ordena
    colunas_final = ["Cliente", "MÊS", "BUDGET", "Importação", "Exportação", "Cabotagem", "Quantidade_iTRACKER",
                     "Aproveitamento de Oportunidade (%)", "Realização do Budget (%)", "Desvio Budget vs Oportunidade (%)",
                     "Target Diário Esperado", "Target Acumulado", "Gap de Realização"]
    df_merged = df_merged[colunas_final].sort_values(by=["Cliente", "MÊS"])
    
    df_merged.to_excel("comparativo_final_atualizado.xlsx", index=False)
    logging.info("Arquivo gerado: comparativo_final_atualizado.xlsx")
    logging.info("Enviando comparativo_final_atualizado.xlsx para o Google Drive...")
    drive_service = authenticate_drive()
    sucesso_upload_final = update_file_in_drive(drive_service, FINAL_ATUALIZADO_FILE_ID, "comparativo_final_atualizado.xlsx")
    if sucesso_upload_final:
        logging.info("comparativo_final_atualizado.xlsx atualizado com sucesso no Google Drive.")
    else:
        logging.error("Falha ao atualizar comparativo_final_atualizado.xlsx no Google Drive.")

#--------------------------------------
# FUNÇÃO PRINCIPAL UNIFICADA
def main():
    logging.info("--------------------------------------------------")
    logging.info("Iniciando Script 1: Processamento Budget vs LogComex e atualização no Google Drive...")
    main_script1()
    logging.info("--------------------------------------------------")
    logging.info("Iniciando Script 2: Processamento iTRACKER e merge completo com a base comparativa...")
    main_script2()
    logging.info("--------------------------------------------------")
    logging.info("Processamento concluído com sucesso.")

if __name__ == "__main__":
    main()
