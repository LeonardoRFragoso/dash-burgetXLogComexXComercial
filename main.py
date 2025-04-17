import subprocess
import logging
import requests
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os.path
import time
import pandas as pd
import os
import unicodedata
import re

# Configuração do logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename=f'execucao_scripts_{datetime.now().strftime("%Y%m%d")}.log'
)

# Configuração do Telegram
TELEGRAM_BOT_TOKEN = "7660740075:AAG0zy6T3QV6pdv2VOwRlxShb0UzVlNwCUk"  # Substitua pelo token do seu bot
TELEGRAM_CHAT_ID = "833732395"  # Substitua pelo Chat ID correto

def enviar_mensagem_telegram(mensagem):
    """Envia uma mensagem para o Telegram."""
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": mensagem,
            "parse_mode": "Markdown"
        }
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            logging.info("Mensagem enviada para o Telegram com sucesso")
        else:
            logging.error(f"Erro ao enviar mensagem para o Telegram: {response.text}")
    except Exception as e:
        logging.error(f"Erro ao tentar enviar mensagem para o Telegram: {e}")

def get_week_dates():
    """Retorna as datas de início e fim da semana atual"""
    hoje = datetime.now()
    dias_desde_segunda = hoje.weekday()
    inicio_semana = hoje - timedelta(days=dias_desde_segunda)
    fim_semana = inicio_semana + timedelta(days=7)
    return inicio_semana, fim_semana

def get_file_names():
    """Gera os nomes dos arquivos sem datas, apenas com seus tipos básicos"""
    return {
        'importacao': 'Importação.xlsx',
        'exportacao': 'Exportação.xlsx',
        'cabotagem': 'Cabotagem.xlsx'
    }

def setup_google_drive():
    """Configura a autenticação com o Google Drive usando conta de serviço"""
    try:
        SCOPES = ['https://www.googleapis.com/auth/drive.file']
        credentials = service_account.Credentials.from_service_account_file(
            'service_account.json',
            scopes=SCOPES
        )
        logging.info("Credenciais da conta de serviço carregadas com sucesso")
        service = build('drive', 'v3', credentials=credentials)
        return service
    except Exception as e:
        logging.error(f"Erro ao configurar autenticação do Google Drive: {e}")
        raise

def update_drive_file(service, file_path, folder_id, mime_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
    """Atualiza ou cria arquivo no Google Drive"""
    try:
        if not os.path.exists(file_path):
            logging.error(f"Arquivo local não encontrado: {file_path}")
            return False

        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)
        logging.info(f"Processando arquivo: {file_name} (Tamanho: {file_size/1024:.2f} KB)")

        if file_size == 0:
            logging.error(f"Arquivo {file_name} está vazio")
            return False

        results = service.files().list(
            q=f"name='{file_name}' and '{folder_id}' in parents and trashed=false",
            fields="files(id, name, size)"
        ).execute()
        files = results.get('files', [])

        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }

        media = MediaFileUpload(
            file_path,
            mimetype=mime_type,
            resumable=True,
            chunksize=1024*1024
        )

        max_retries = 3
        retry_count = 0

        while retry_count < max_retries:
            try:
                if files:
                    file_id = files[0]['id']
                    logging.info(f"Atualizando arquivo existente no Drive: {file_name}")
                    service.files().update(
                        fileId=file_id,
                        media_body=media,
                        fields='id'
                    ).execute()
                    logging.info(f"Arquivo atualizado com sucesso: {file_name}")
                else:
                    logging.info(f"Criando novo arquivo no Drive: {file_name}")
                    file = service.files().create(
                        body=file_metadata,
                        media_body=media,
                        fields='id'
                    ).execute()
                    logging.info(f"Novo arquivo criado com sucesso: {file_name} (ID: {file['id']})")
                return True
            except Exception as e:
                retry_count += 1
                if retry_count < max_retries:
                    wait_time = 5 * retry_count
                    logging.warning(f"Tentativa {retry_count} falhou para {file_name}. Aguardando {wait_time}s antes de tentar novamente. Erro: {str(e)}")
                    time.sleep(wait_time)
                else:
                    logging.error(f"Falha após {max_retries} tentativas para {file_name}: {str(e)}")
                    return False
    except Exception as e:
        logging.error(f"Erro ao processar arquivo {file_path}: {str(e)}")
        return False

def sincronizar_com_drive():
    """Sincroniza os arquivos Excel com o Google Drive, incluindo a planilha consolidada."""
    try:
        logging.info("Iniciando sincronização com Google Drive")
        drive_service = setup_google_drive()
        folder_id = "1WAPVG7dbzkve9Vx2NvQ7kbDfnchnDRMi"
        
        # Obter os nomes de arquivos com as datas corretas
        nomes_arquivos = get_file_names()
        
        arquivos_para_sincronizar = [
            #nomes_arquivos['importacao'],
            #nomes_arquivos['exportacao'],
            #nomes_arquivos['cabotagem'],
            'Dados_Consolidados.xlsx'  # Nome fixo para o arquivo consolidado
        ]

        arquivos_sincronizados = 0
        for arquivo in arquivos_para_sincronizar:
            caminho_completo = os.path.abspath(arquivo)
            logging.info(f"Verificando arquivo para sincronização: {caminho_completo}")
            
            if os.path.exists(caminho_completo):
                try:
                    # Verificar se o arquivo não está vazio
                    if os.path.getsize(caminho_completo) > 0:
                        resultado = update_drive_file(drive_service, caminho_completo, folder_id)
                        if resultado:
                            arquivos_sincronizados += 1
                        else:
                            logging.error(f"Falha ao sincronizar arquivo {arquivo}")
                    else:
                        logging.error(f"Arquivo {arquivo} está vazio")
                except Exception as e:
                    logging.error(f"Erro ao processar arquivo {arquivo}: {e}")
            else:
                logging.warning(f"Arquivo {arquivo} não encontrado em {caminho_completo}")
        
        logging.info(f"Sincronização com Google Drive concluída. {arquivos_sincronizados} de {len(arquivos_para_sincronizar)} arquivos sincronizados.")
        return arquivos_sincronizados > 0
    except Exception as e:
        logging.error(f"Erro na sincronização com Google Drive: {e}")
        return False

def criar_planilha_consolidada():
    """
    Gera a planilha Dados_Consolidados.xlsx unindo Importação, Exportação e Cabotagem,
    lidando com diferentes estruturas de colunas entre as planilhas.
    """
    # Obtendo nomes dos arquivos gerados na semana
    arquivos = get_file_names()
    
    planilhas = {
        "Importação": arquivos['importacao'],
        "Exportação": arquivos['exportacao'],
        "Cabotagem": arquivos['cabotagem']
    }

    # Verificar se pelo menos um arquivo existe antes de continuar
    arquivos_existentes = [arquivo for categoria, arquivo in planilhas.items() 
                         if os.path.exists(arquivo) and os.path.getsize(arquivo) > 0]
    
    if not arquivos_existentes:
        logging.error("Nenhum arquivo válido encontrado para consolidação. Verifique se os scripts geraram as planilhas corretamente.")
        return False

    logging.info(f"Arquivos encontrados para consolidação: {', '.join(arquivos_existentes)}")

    # Lista para armazenar os DataFrames
    dfs = []
    # Set para coletar todas as colunas únicas
    todas_colunas = set()
    
    # Primeiro loop: coletar todas as colunas possíveis
    for categoria, arquivo in planilhas.items():
        if os.path.exists(arquivo) and os.path.getsize(arquivo) > 0:
            try:
                df = pd.read_excel(arquivo)
                # Adicionar todas as colunas ao set (exceto a categoria que será adicionada depois)
                todas_colunas.update(df.columns)
                logging.info(f"Colunas encontradas em {categoria}: {', '.join(df.columns)}")
            except Exception as e:
                logging.error(f"Erro ao analisar colunas de {arquivo}: {e}")
    
    # Converter o set para lista e ordenar para manter consistência
    colunas_padronizadas = sorted(list(todas_colunas))
    logging.info(f"Total de colunas únicas encontradas: {len(colunas_padronizadas)}")
    
    # Segundo loop: carregar e padronizar cada DataFrame
    for categoria, arquivo in planilhas.items():
        if os.path.exists(arquivo) and os.path.getsize(arquivo) > 0:
            try:
                df = pd.read_excel(arquivo)
                
                # Criar DataFrame vazio com todas as colunas
                df_padronizado = pd.DataFrame(columns=colunas_padronizadas)
                
                # Copiar dados existentes
                for coluna in colunas_padronizadas:
                    if coluna in df.columns:
                        df_padronizado[coluna] = df[coluna]
                    else:
                        # Preencher com NaN quando a coluna não existe
                        df_padronizado[coluna] = pd.NA
                
                # Adicionar a coluna de categoria no início
                df_padronizado.insert(0, "Categoria", categoria)
                
                dfs.append(df_padronizado)
                logging.info(f"DataFrame de {categoria} padronizado com sucesso")
                
            except Exception as e:
                logging.error(f"Erro ao processar {arquivo}: {e}")
    
    if dfs:
        try:
            # Concatenar todos os DataFrames
            df_consolidado = pd.concat(dfs, ignore_index=True)
            
            # Ordenar as colunas mantendo "Categoria" como primeira
            colunas_ordenadas = ["Categoria"] + [col for col in colunas_padronizadas if col != "Categoria"]
            df_consolidado = df_consolidado[colunas_ordenadas]
            
            # Salvar o arquivo consolidado
            arquivo_consolidado = "Dados_Consolidados.xlsx"
            df_consolidado.to_excel(arquivo_consolidado, index=False)
            
            # Verificar se o arquivo foi criado corretamente
            if os.path.exists(arquivo_consolidado) and os.path.getsize(arquivo_consolidado) > 0:
                logging.info(f"Planilha {arquivo_consolidado} criada com sucesso (Tamanho: {os.path.getsize(arquivo_consolidado)/1024:.2f} KB)")
                
                # Log com estatísticas básicas
                logging.info("Estatísticas da consolidação:")
                logging.info(f"- Total de registros: {len(df_consolidado)}")
                logging.info(f"- Total de colunas: {len(df_consolidado.columns)}")
                for categoria in df_consolidado['Categoria'].unique():
                    count = len(df_consolidado[df_consolidado['Categoria'] == categoria])
                    logging.info(f"- Registros de {categoria}: {count}")
                return True
            else:
                logging.error(f"Falha ao verificar o arquivo consolidado: {arquivo_consolidado}")
                return False
                
        except Exception as e:
            logging.error(f"Erro ao criar arquivo consolidado: {e}")
            return False
    else:
        logging.warning("Nenhuma planilha foi carregada para consolidar.")
        return False

def normalize_text(text):
    """Normaliza o texto removendo acentos e convertendo para minúsculas."""
    return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8').lower()

def inserir_info_clientes():
    """
    Lê o arquivo de clientes e insere na planilha 'Dados_Consolidados.xlsx' duas novas colunas:
    - 'Clientes Encontrados': com todos os clientes encontrados na linha, separados por vírgula.
    - 'Colunas Encontradas': informando em qual(is) coluna(s) o cliente foi localizado.
    Se nenhum cliente for identificado, as colunas permanecem vazias.
    
    Alterações realizadas:
    1. A busca é feita apenas nas colunas específicas:
       AGENTE DE CARGA, AGENTE INTERNACIONAL, ARMADOR, CONSIGNATARIO FINAL,
       CONSIGNATÁRIO, CONSOLIDADOR, DESTINATÁRIO, NOME EXPORTADOR, NOME IMPORTADOR, REMETENTE.
    2. Utiliza-se expressão regular com delimitadores de palavra para evitar correspondências parciais indesejadas.
    """
    try:
        arquivo_consolidado = "Dados_Consolidados.xlsx"
        
        # Verificar se o arquivo consolidado existe
        if not os.path.exists(arquivo_consolidado):
            logging.error(f"Arquivo {arquivo_consolidado} não encontrado para inserção de informações de clientes.")
            return False
            
        if os.path.getsize(arquivo_consolidado) == 0:
            logging.error(f"Arquivo {arquivo_consolidado} está vazio.")
            return False
            
        # Carrega a planilha consolidada
        df = pd.read_excel(arquivo_consolidado)
        logging.info(f"Planilha {arquivo_consolidado} carregada para análise dos clientes. Total de registros: {len(df)}")
        
        # Lista de colunas onde buscar os nomes dos clientes
        search_columns = [
            "AGENTE DE CARGA",
            "AGENTE INTERNACIONAL",
            "ARMADOR",
            "CONSIGNATARIO FINAL",
            "CONSIGNATÁRIO",
            "CONSOLIDADOR",
            "DESTINATÁRIO",
            "NOME EXPORTADOR",
            "NOME IMPORTADOR",
            "REMETENTE"
        ]
        
        # Verificar quais colunas de busca existem no DataFrame
        colunas_existentes = [col for col in search_columns if col in df.columns]
        logging.info(f"Colunas de busca encontradas no DataFrame: {', '.join(colunas_existentes)}")
        
        if not colunas_existentes:
            logging.warning("Nenhuma coluna de busca encontrada no DataFrame. Não será possível identificar clientes.")
            return False
        
        # Caminho absoluto para o arquivo de clientes
        clientes_path = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Projeto-Comercial\clientes.txt"
        
        # Verificar se o arquivo de clientes existe
        if not os.path.exists(clientes_path):
            logging.error(f"Arquivo de clientes não encontrado: {clientes_path}")
            return False
            
        with open(clientes_path, "r", encoding="utf-8") as f:
            # Remover linhas vazias e espaços extras
            clientes = [linha.strip() for linha in f if linha.strip()]
        
        if not clientes:
            logging.warning("Nenhum cliente encontrado no arquivo de clientes.")
            return False
            
        # Cria uma lista de tuplas (cliente, cliente_normalizado)
        clientes_norm = [(cliente, normalize_text(cliente)) for cliente in clientes]
        logging.info(f"{len(clientes_norm)} clientes carregados do arquivo.")
        
        # Listas para armazenar os resultados para cada linha da planilha
        lista_clientes_encontrados = []
        lista_colunas_encontradas = []
        
        # Percorre cada linha da planilha
        for idx, row in df.iterrows():
            matches = {}
            # Itera somente pelas colunas definidas em colunas_existentes
            for col in colunas_existentes:
                cell_value = str(row[col]) if pd.notna(row[col]) else ""
                cell_norm = normalize_text(cell_value)
                # Para cada cliente, utiliza regex com delimitadores de palavra para busca
                for cliente, cliente_norm in clientes_norm:
                    pattern = r'\b' + re.escape(cliente_norm) + r'\b'
                    if re.search(pattern, cell_norm):
                        if cliente not in matches:
                            matches[cliente] = set()
                        matches[cliente].add(col)
            if matches:
                # Lista de clientes encontrados separados por vírgula
                clientes_encontrados = ", ".join(sorted(matches.keys()))
                # Lista única das colunas onde os clientes foram encontrados
                colunas_encontradas = ", ".join(sorted(set.union(*matches.values())))
            else:
                clientes_encontrados = ""
                colunas_encontradas = ""
            lista_clientes_encontrados.append(clientes_encontrados)
            lista_colunas_encontradas.append(colunas_encontradas)
        
        # Adiciona as novas colunas à planilha
        df["Clientes Encontrados"] = lista_clientes_encontrados
        df["Colunas Encontradas"] = lista_colunas_encontradas
        
        # Conta o número de linhas com clientes encontrados
        linhas_com_clientes = sum(1 for cliente in lista_clientes_encontrados if cliente)
        logging.info(f"Clientes encontrados em {linhas_com_clientes} de {len(df)} registros.")
        
        # Salva a planilha atualizada
        df.to_excel(arquivo_consolidado, index=False)
        
        # Verificar se o arquivo foi salvo corretamente
        if os.path.exists(arquivo_consolidado) and os.path.getsize(arquivo_consolidado) > 0:
            logging.info(f"Informações dos clientes inseridas com sucesso na planilha {arquivo_consolidado}")
            logging.info(f"Tamanho final do arquivo: {os.path.getsize(arquivo_consolidado)/1024:.2f} KB")
            return True
        else:
            logging.error(f"Erro ao salvar arquivo após inserção de informações dos clientes: {arquivo_consolidado}")
            return False
            
    except Exception as e:
        logging.error(f"Erro ao inserir informações dos clientes: {e}")
        return False

def executar_script(nome_script, max_tentativas=3, tempo_espera=30):
    tentativa = 1
    while tentativa <= max_tentativas:
        try:
            logging.info(f"Iniciando execução do script: {nome_script} (Tentativa {tentativa}/{max_tentativas})")
            processo = subprocess.Popen(
                ['python', nome_script],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1
            )
            while True:
                linha = processo.stdout.readline()
                if not linha and processo.poll() is not None:
                    break
                if linha:
                    print(linha.strip())
                    logging.info(linha.strip())
            retorno = processo.wait()
            if retorno == 0:
                logging.info(f"Script {nome_script} finalizado com sucesso")
                return True
            else:
                logging.error(f"Script {nome_script} falhou com código de retorno {retorno}")
                if tentativa < max_tentativas:
                    time.sleep(tempo_espera)
                tentativa += 1
        except Exception as e:
            logging.error(f"Erro ao executar {nome_script}: {e}")
            if tentativa < max_tentativas:
                time.sleep(tempo_espera)
            tentativa += 1
    return False

def main():
    scripts = [
        os.path.join('mes atual', 'app_exportacao.py'),
        os.path.join('mes atual', 'app_importacao.py'),
        os.path.join('mes atual', 'app_cabotagem.py')
    ]

    logging.info("Iniciando execução sequencial dos scripts")
    scripts_executados = []
    scripts_falhos = []
    max_ciclos = 30

    for script in scripts:
        if executar_script(script):
            scripts_executados.append(script)
            time.sleep(30)
        else:
            scripts_falhos.append(script)

    if scripts_falhos:
        ciclo = 1
        while scripts_falhos and ciclo < max_ciclos:
            time.sleep(60)
            ainda_falhos = []
            for script in scripts_falhos:
                if executar_script(script, max_tentativas=2):
                    scripts_executados.append(script)
                else:
                    ainda_falhos.append(script)
            scripts_falhos = ainda_falhos
            ciclo += 1

    # Verifica se ao menos um script foi executado com sucesso
    if not scripts_executados:
        logging.error("Nenhum script foi executado com sucesso. Não é possível gerar a planilha consolidada.")
        enviar_mensagem_telegram("\ud83d\udea8 *Falha na execução de todos os scripts!*\n\nNenhum script foi executado com sucesso.")
        return False

    # Cria a planilha consolidada após execução dos scripts
    planilha_consolidada_criada = criar_planilha_consolidada()
    if not planilha_consolidada_criada:
        logging.error("Falha ao criar a planilha consolidada.")
        enviar_mensagem_telegram("\ud83d\udea8 *Falha na geração da planilha consolidada!*\n\nOs scripts foram executados, mas não foi possível gerar a planilha consolidada.")
        return False
    
    # Insere as informações dos clientes na planilha consolidada
    clientes_inseridos = inserir_info_clientes()
    if not clientes_inseridos:
        logging.error("Falha ao inserir informações dos clientes na planilha consolidada.")
        enviar_mensagem_telegram("\u26a0\ufe0f *Aviso: Falha ao inserir informações de clientes!*\n\nA planilha consolidada foi gerada, mas não foi possível inserir as informações dos clientes.")
    
    # Verifica quais arquivos estão disponíveis para sincronização
    arquivos = get_file_names()
    arquivos_disponiveis = []
    
    for tipo, nome in arquivos.items():
        if os.path.exists(nome) and os.path.getsize(nome) > 0:
            arquivos_disponiveis.append(nome)
    
    if os.path.exists("Dados_Consolidados.xlsx") and os.path.getsize("Dados_Consolidados.xlsx") > 0:
        arquivos_disponiveis.append("Dados_Consolidados.xlsx")
    
    if not arquivos_disponiveis:
        logging.error("Nenhum arquivo disponível para sincronização com o Google Drive.")
        enviar_mensagem_telegram("\ud83d\udea8 *Falha: Nenhum arquivo disponível para sincronização!*\n\nOs scripts foram executados, mas não foi possível encontrar os arquivos para sincronização.")
        return False
    
    logging.info(f"Arquivos disponíveis para sincronização: {', '.join(arquivos_disponiveis)}")

    # Tenta sincronizar com o Google Drive
    sincronizacao_concluida = False
    if arquivos_disponiveis:
        time.sleep(30)
        for tentativa in range(3):
            try:
                if sincronizar_com_drive():
                    sincronizacao_concluida = True
                    break
                else:
                    if tentativa < 2:
                        logging.warning(f"Tentativa {tentativa + 1} de sincronização falhou. Aguardando 30 segundos para nova tentativa.")
                        time.sleep(30)
            except Exception as e:
                logging.error(f"Erro na tentativa {tentativa + 1} de sincronização: {e}")
                if tentativa < 2:
                    time.sleep(30)

    # Monta a mensagem de status final
    if scripts_falhos:
        mensagem = (
            f"\ud83d\udea8 *Execução com falhas!*\n\n"
            f"Scripts com falha: {', '.join(scripts_falhos)}\n"
            f"Scripts executados: {', '.join(scripts_executados)}\n"
        )
        if sincronizacao_concluida:
            mensagem += f"Arquivos sincronizados com sucesso: {', '.join(arquivos_disponiveis)}"
        else:
            mensagem += f"Falha na sincronização com o Google Drive."
    elif sincronizacao_concluida:
        mensagem = (
            f"\u2705 *Execução concluída com sucesso!*\n\n"
            f"Todos os scripts foram executados e os seguintes arquivos foram sincronizados no Google Drive:\n"
            f"{', '.join(arquivos_disponiveis)}"
        )
    else:
        mensagem = (
            f"\u26a0\ufe0f *Execução parcialmente concluída!*\n\n"
            f"Scripts executados: {', '.join(scripts_executados)}\n"
            f"Falha durante a sincronização com o Google Drive."
        )

    enviar_mensagem_telegram(mensagem)
    return len(scripts_falhos) == 0 and sincronizacao_concluida

if __name__ == "__main__":
    main()
