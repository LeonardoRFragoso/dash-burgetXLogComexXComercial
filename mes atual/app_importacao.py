import time
import requests
import logging
import re
import os
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Parâmetros para formatação e cabeçalhos esperados (fluxo de importação)
EXPECTED_HEADERS = [
    'ID', 'EMBARQUE', 'CONSIGNATARIO FINAL', 'CONSOLIDADOR', 'CONSIGNATÁRIO',
    'TIPO CARGA', 'ETS', 'ETA', 'TRANSIT-TIME', 'SAÍDA PORTO', 'TEMPO PORTO',
    'PAÍS ORIGEM', 'PAÍS DE EMBARQUE', 'VIAGEM', 'PORTO EMBARQUE', 'PORTO DESCARGA',
    'TERMINAL DESCARGA', 'UF CONSIGNATÁRIO', 'CNPJ CONSIGNATÁRIO',
    'QTDE CONSIGNATÁRIO FINAL', 'ATIVIDADE DO CONSIGNATARIO',
    'CIDADE DO CONSIGNATÁRIO', 'EMAIL', 'TELEFONE', 'PAGAMENTO', 'CARGA PERIGOSA',
    'TIPO CONTAINER', 'MERCADORIA', 'QUANTIDADE VEICULOS', 'QTDE CONTAINER',
    'TEUS', 'C20', 'C40', 'VOLUMES', 'PESO BRUTO', 'ARMAZEM DESTINO',
    'TRADE LANE', 'NVOCC', 'AGENTE DE CARGA', 'NAVIO', 'PAÍS DE PROCEDÊNCIA',
    'PORTO ORIGEM', 'PORTO DESTINO', 'NOTIFICADO', 'NOME EXPORTADOR',
    'PORTO DESCARGA COM CÓDIGO', 'ARMADOR', 'AGENTE INTERNACIONAL', 'HS CODE',
    'CONTAINER PARCIAL', 'PORTO ORIGEM COM CÓDIGO', 'PORTO DESTINO COM CÓDIGO',
    'CNPJ AGENTE DE CARGA', 'PROVÁVEL LOCAL DE LIBERAÇÃO', 'ANO/MÊS'
]
NUMERIC_COLUMNS = [
    'TRANSIT-TIME', 'TEMPO PORTO', 'QTDE CONSIGNATÁRIO FINAL',
    'QUANTIDADE VEICULOS', 'QTDE CONTAINER', 'TEUS', 'C20', 'C40',
    'VOLUMES', 'PESO BRUTO'
]
DATE_COLUMNS = ['ETS', 'ETA', 'SAÍDA PORTO']

def get_date_range_7_days():
    hoje = datetime.now()
    inicio = hoje - timedelta(days=7)
    fim = hoje + timedelta(days=30)
    return inicio.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y")

def get_date_range_30_dias_antes_depois():
    hoje = datetime.now()
    inicio = hoje - timedelta(days=30)
    fim = hoje + timedelta(days=30)
    return inicio.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y")

def get_date_range_30_days():
    hoje = datetime.now()
    inicio = hoje - timedelta(days=30)
    fim = hoje + timedelta(days=30)
    return inicio.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y")

# NOVA FUNÇÃO: Retorna o intervalo do mês atual completo
def get_date_range_current_month():
    hoje = datetime.now()
    inicio = datetime(hoje.year, hoje.month, 1)
    # Calcula o último dia do mês atual
    if hoje.month == 12:
        next_month = datetime(hoje.year + 1, 1, 1)
    else:
        next_month = datetime(hoje.year, hoje.month + 1, 1)
    fim = next_month - timedelta(days=1)
    return inicio.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y")

def format_numeric_value(value):
    try:
        if pd.isna(value) or value == '' or value == 'nan':
            return '0,00'
        clean_value = re.sub(r'[^\d.,]', '', str(value))
        clean_value = clean_value.replace(',', '.')
        float_value = float(clean_value)
        return f"{float_value:.2f}".replace('.', ',')
    except Exception as e:
        logger.warning(f"Erro ao formatar valor numérico '{value}': {str(e)}")
        return '0,00'

def scroll_table_fully(driver):
    row_xpath = "//*[@id='pdf-export-container']/div[2]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[4]/div[1]/div[2]/div/div"
    actions = ActionChains(driver)
    try:
        max_attempts = 20
        attempt = 0
        last_row_count = 0
        while attempt < max_attempts:
            rows = driver.find_elements(By.XPATH, row_xpath)
            num_rows = len(rows)
            logger.info(f"{num_rows} linhas atualmente visíveis.")
            if num_rows >= 30:
                logger.info("30 linhas capturadas, rolagem completa.")
                break
            if num_rows == last_row_count:
                actions.scroll_by_amount(0, 30).perform()
            else:
                actions.move_to_element(rows[-1]).perform()
            last_row_count = num_rows
            time.sleep(1)
            attempt += 1
        if attempt == max_attempts:
            logger.warning("Máximo de tentativas de rolagem atingido. Linhas ainda incompletas.")
    except Exception as e:
        logger.warning(f"Erro durante a rolagem: {str(e)}")

def wait_for_table_load(driver, timeout=15):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ag-center-cols-container"))
    )
    logger.info("Tabela completamente carregada.")

def process_data_frame(all_data, header_texts):
    try:
        logger.info("Iniciando processamento do DataFrame.")
        df = pd.DataFrame(all_data, columns=header_texts)
        df['DATA CONSULTA'] = datetime.now().strftime('%d/%m/%Y')
        if 'CONSIGNATARIO FINAL' in df.columns:
            df['NOME IMPORTADOR'] = df['CONSIGNATARIO FINAL']
        else:
            df['NOME IMPORTADOR'] = ''
            logger.warning("Coluna 'CONSIGNATARIO FINAL' não encontrada. 'NOME IMPORTADOR' será criada vazia.")
        df = df.reindex(columns=EXPECTED_HEADERS + ['DATA CONSULTA', 'NOME IMPORTADOR'])
        df = df.drop_duplicates()
        df = df.dropna(how='all')
        logger.info(f"DataFrame após limpeza: {df.shape}")
        if 'ID' in df.columns:
            df['ID'] = df['ID'].astype(str).str.replace('.', '', regex=True).str.strip()
            logger.info("Coluna ID tratada como string, removendo separadores de milhar.")
        df = corrigir_tipos_dados(df)
        for col in NUMERIC_COLUMNS:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: format_numeric_value(str(x).strip()))
        logger.info(f"Processamento do DataFrame concluído. Shape final: {df.shape}")
        return df
    except Exception as e:
        logger.error(f"Erro no processamento do DataFrame: {str(e)}")
        raise

def corrigir_tipos_dados(df):
    try:
        logger.info(f"Iniciando correção de tipos de dados. Shape inicial: {df.shape}")
        campos_numericos = [
            'TRANSIT-TIME', 'TEMPO PORTO', 'QTDE CONSIGNATÁRIO FINAL',
            'QUANTIDADE VEICULOS', 'QTDE CONTAINER', 'TEUS', 'C20', 'C40',
            'VOLUMES', 'PESO BRUTO'
        ]
        if 'ID' in df.columns:
            df['ID'] = df['ID'].astype(str).str.replace('.', '', regex=True).str.strip()
            logger.info("Coluna ID tratada como string, removendo separadores de milhar.")
        def corrigir_numero(valor):
            try:
                if pd.isna(valor) or valor == '' or valor is None:
                    return 0.0
                return float(str(valor).replace('.', '').replace(',', '.'))
            except ValueError:
                return 0.0
        for campo in campos_numericos:
            if campo in df.columns:
                try:
                    df[campo] = df[campo].astype(str).apply(corrigir_numero)
                    df[campo] = df[campo].fillna(0)
                    logger.info(f"Campo {campo} convertido para numérico")
                except Exception as e:
                    logger.warning(f"Erro ao converter campo {campo}: {str(e)}")
        campos_data = ['ETS', 'ETA', 'SAÍDA PORTO']
        for campo in campos_data:
            if campo in df.columns:
                try:
                    df[campo] = pd.to_datetime(df[campo], format='%Y-%m-%d', errors='coerce')
                    df[campo] = df[campo].dt.strftime('%d/%m/%Y').fillna('').str.strip()
                    logger.info(f"Campo {campo} convertido para data")
                except Exception as e:
                    logger.warning(f"Erro ao converter campo {campo}: {str(e)}")
        if 'ANO/MÊS' in df.columns:
            def converter_anomes(valor):
                try:
                    if pd.isna(valor) or valor == '' or valor is None:
                        return None
                    if isinstance(valor, str):
                        valor = valor.strip().replace(' ', '')
                        if '.' in valor:
                            ano, mes = valor.split('.')
                            return int(ano + mes.zfill(2))
                        return int(valor)
                    return int(valor)
                except:
                    return None
            df['ANO/MÊS'] = df['ANO/MÊS'].apply(converter_anomes)
            ano_mes_atual = int(datetime.now().strftime('%Y%m'))
            df['ANO/MÊS'] = df['ANO/MÊS'].fillna(ano_mes_atual).astype(int)
            logger.info("Coluna ANO/MÊS processada e convertida para inteiro")
        if 'DATA CONSULTA' in df.columns:
            df['DATA CONSULTA'] = df['DATA CONSULTA'].astype(str)
            logger.info("Coluna DATA CONSULTA convertida para string")
        logger.info(f"Correção de tipos concluída. Shape final: {df.shape}")
        return df
    except Exception as e:
        logger.error(f"Erro no processamento do DataFrame: {str(e)}")
        logger.error(f"Colunas do DataFrame: {df.columns.tolist()}")
        raise

def get_total_pages(driver):
    try:
        pagination_xpath = "//span[@class='ag-paging-description']"
        pagination_text = driver.find_element(By.XPATH, pagination_xpath).text
        total_pages = int(re.search(r"de (\d+)", pagination_text).group(1))
        logger.info(f"Total de páginas detectadas: {total_pages}")
        return total_pages
    except Exception as e:
        logger.warning(f"Falha ao obter o total de páginas: {str(e)}")
        return 1

def navigate_to_next_page(driver, wait, current_page):
    try:
        next_button_container = driver.find_element(By.XPATH, "//*[@id='ag-42']")
        next_button = next_button_container.find_element(By.CLASS_NAME, "ag-icon-next")
        if next_button.is_displayed() and next_button.is_enabled():
            next_button.click()
            logger.info(f"Processando página {current_page + 1}")
            time.sleep(2)
            return True
        else:
            logger.info("Última página alcançada.")
            return False
    except Exception as e:
        logger.error(f"Erro ao tentar clicar na próxima página: {e}")
        return False

def extract_table_data(driver, wait, header_texts, max_retries=3):
    for attempt in range(max_retries):
        try:
            script = """
            let rows = document.querySelectorAll('#pdf-export-container .ag-center-cols-container .ag-row');
            let rowData = [];
            rows.forEach(row => {
                let cells = row.querySelectorAll('.ag-cell');
                let data = Array.from(cells).map(cell => cell.textContent.trim());
                rowData.push(data);
            });
            return rowData;
            """
            rows = driver.execute_script(script)
            if rows:
                logger.info(f"{len(rows)} linhas capturadas.")
                page_data = []
                for row in rows:
                    while len(row) < len(header_texts):
                        row.append('')
                    if len(row) == len(header_texts):
                        page_data.append(row)
                    else:
                        logger.warning(f"Linha inconsistente, preenchida com valores vazios: {row}")
                return page_data
        except Exception as e:
            logger.error(f"Erro na tentativa {attempt + 1}: {str(e)}")
            driver.save_screenshot(f"erro_extracao_{attempt}.png")
            time.sleep(2)
    logger.error("Todas as tentativas de extração falharam")
    return []

def acessar_shipment_intel(driver, wait):
    """
    Aguarda e clica no botão "Acessar" do card Shipment Intel, utilizando o XPath informado.
    """
    try:
        shipment_intel_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div/main/div/div/div[1]/div[1]/div[3]/div/div/div/div[1]/div[2]/button'))
        )
        shipment_intel_button.click()
        logger.info("Clicado no botão 'Acessar' do card Shipment Intel.")
        time.sleep(3)
    except Exception as e:
        logger.error("Erro ao acessar Shipment Intel: " + str(e))
        driver.save_screenshot("erro_shipment_intel.png")
        raise

def navegar_para_secao(driver, wait):
    try:
        # Clica no botão do menu lateral
        menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LSTopBreadcrumbMenuActivator"]/div[1]')))
        ActionChains(driver).move_to_element(menu_button).click().perform()
        logger.info("Menu lateral aberto.")
        time.sleep(2)

        # Clica em "Importação Marítima"
        importacao_xpath = '//*[@id="app"]/div[2]/div/div/div/div[1]/div[2]/ul/li[4]'
        importacao = wait.until(EC.element_to_be_clickable((By.XPATH, importacao_xpath)))
        ActionChains(driver).move_to_element(importacao).click().perform()
        logger.info("Importação Marítima clicada.")
        time.sleep(2)

        # Clica em "Brasil"
        brasil_xpath = '//*[@id="app"]/div[2]/div/div/div/div[2]/div[2]/ul/li'
        brasil = wait.until(EC.element_to_be_clickable((By.XPATH, brasil_xpath)))
        ActionChains(driver).move_to_element(brasil).click().perform()
        logger.info("Brasil selecionado.")
        time.sleep(2)
    except Exception as e:
        logger.error(f"Erro ao navegar para seção de importação: {e}")
        driver.save_screenshot("erro_navegacao_importacao.png")
        raise

def unificar_valor(valor):
    if pd.isna(valor) or valor == '':
        return ''
    return re.sub(r'\s+', ' ', valor.strip()).upper()

def adicionar_aba_resumo(df, file_name):
    try:
        resumo_df = df[['CONSIGNATÁRIO', 'PORTO DESTINO', 'ETA', 'QTDE CONTAINER']].copy()
        resumo_df['CONSIGNATÁRIO'] = resumo_df['CONSIGNATÁRIO'].apply(unificar_valor)
        resumo_df['PORTO DESTINO'] = resumo_df['PORTO DESTINO'].apply(unificar_valor)
        resumo_df['ETA'] = pd.to_datetime(resumo_df['ETA'], errors='coerce').dt.strftime('%d/%m/%Y').str.strip()
        resumo_df['QTDE CONTAINER'] = (resumo_df['QTDE CONTAINER']
            .astype(str)
            .str.replace(',', '.')
            .pipe(pd.to_numeric, errors='coerce')
            .fillna(0)
        )
        resumo_agrupado = resumo_df.groupby(
            ['CONSIGNATÁRIO', 'PORTO DESTINO', 'ETA'],
            as_index=False
        ).agg({'QTDE CONTAINER': 'sum'})
        resumo_agrupado = resumo_agrupado.sort_values(['CONSIGNATÁRIO', 'PORTO DESTINO', 'ETA']).reset_index(drop=True)
        resumo_agrupado['QTDE CONTAINER'] = resumo_agrupado['QTDE CONTAINER'].astype(int)
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            resumo_agrupado.to_excel(writer, sheet_name='Resumo', index=False)
        logger.info(f"Aba 'Resumo' gerada com sucesso no arquivo: {file_name}")
        return resumo_agrupado
    except Exception as e:
        logger.error(f"Erro ao criar aba 'Resumo': {str(e)}")
        raise

def ajustar_zoom(driver, zoom_level=80):
    try:
        driver.execute_script(f"document.body.style.zoom = '{zoom_level}%'")
        logger.info(f"Zoom ajustado para {zoom_level}%")
    except Exception as e:
        logger.warning(f"Não foi possível ajustar o zoom: {e}")

def salvar_arquivo_excel(df, nome_arquivo):
    nome_backup = nome_arquivo.replace('.xlsx', '_backup.xlsx')

    def verificar_arquivo_em_uso(arquivo):
        try:
            with open(arquivo, 'a'):
                return False
        except IOError:
            return True

    if os.path.exists(nome_arquivo):
        try:
            if os.path.exists(nome_backup):
                os.remove(nome_backup)
            os.rename(nome_arquivo, nome_backup)
            logger.info(f"Backup criado: {nome_backup}")
        except Exception as e:
            logger.warning(f"Não foi possível criar backup: {e}")

    def limpar_dados(df):
        for coluna in df.columns:
            if df[coluna].dtype == 'object':
                df[coluna] = df[coluna].fillna('')
            else:
                df[coluna] = df[coluna].fillna(0)
        return df

    df = limpar_dados(df)

    max_tentativas = 3
    tentativa = 1
    while tentativa <= max_tentativas:
        try:
            if verificar_arquivo_em_uso(nome_arquivo):
                logger.warning(f"Arquivo {nome_arquivo} está em uso. Tentativa {tentativa}/{max_tentativas}")
                time.sleep(5)
                tentativa += 1
                continue

            with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Logcomex', index=False)
                logger.info(f"Aba Logcomex exportada com sucesso: {len(df)} registros")

            if os.path.exists(nome_arquivo) and os.path.getsize(nome_arquivo) > 0:
                logger.info(f"Arquivo {nome_arquivo} salvo com sucesso")

                try:
                    adicionar_aba_resumo(df, nome_arquivo)
                    logger.info("Aba de resumo adicionada com sucesso")
                except Exception as e:
                    logger.error(f"Erro ao adicionar aba de resumo: {e}")

                if os.path.exists(nome_backup):
                    os.remove(nome_backup)
                    logger.info("Arquivo de backup removido")

                break
            else:
                raise Exception("Arquivo não foi salvo corretamente")

        except Exception as e:
            logger.error(f"Erro ao salvar arquivo (tentativa {tentativa}): {e}")
            if tentativa == max_tentativas:
                if os.path.exists(nome_backup):
                    if os.path.exists(nome_arquivo):
                        os.remove(nome_arquivo)
                    os.rename(nome_backup, nome_arquivo)
                    logger.info("Backup restaurado após falha")
                raise
            tentativa += 1
            time.sleep(5)

def slow_typing(element, text, delay=0.1):
    for char in text:
        element.send_keys(char)
        time.sleep(delay)

def solve_recaptcha(driver, api_key):
    try:
        recaptcha_button = driver.find_element(By.ID, "signIn__container__card__form__signInButton")
        site_key = recaptcha_button.get_attribute("data-sitekey")
        current_url = driver.current_url
        logger.info(f"Site key encontrada: {site_key}")
    except Exception as e:
        logger.error(f"Não foi possível localizar o botão com reCAPTCHA: {e}")
        raise

    s = requests.Session()
    payload = {
        'key': api_key,
        'method': 'userrecaptcha',
        'googlekey': site_key,
        'pageurl': current_url,
        'json': 1
    }
    response = s.post("http://2captcha.com/in.php", data=payload).json()
    if response['status'] != 1:
        raise Exception("Erro no 2Captcha: " + response['request'])
    captcha_id = response['request']
    logger.info(f"Captcha enviado para 2Captcha, ID: {captcha_id}")

    time.sleep(20)
    captcha_token = None
    for attempt in range(20):
        res = s.get("http://2captcha.com/res.php", params={
            'key': api_key,
            'action': 'get',
            'id': captcha_id,
            'json': 1
        }).json()
        if res['status'] == 1:
            captcha_token = res['request']
            logger.info("Captcha resolvido com sucesso.")
            break
        else:
            logger.info(f"Aguardando solução do captcha... tentativa {attempt + 1}")
            time.sleep(5)
    if not captcha_token:
        raise Exception("Tempo excedido para resolução do captcha.")

    try:
        driver.execute_script("document.getElementById('g-recaptcha-response').style.display='block';")
        driver.execute_script("document.getElementById('g-recaptcha-response').value = arguments[0];", captcha_token)
        driver.execute_script(
            "var event = new Event('input', { bubbles: true });"
            "document.getElementById('g-recaptcha-response').dispatchEvent(event);",
            captcha_token
        )
        logger.info("Token injetado no g-recaptcha-response.")
        callback = driver.execute_script("return document.getElementById('signIn__container__card__form__signInButton').getAttribute('data-callback');")
        if callback:
            driver.execute_script("if(window[arguments[0]]) { window[arguments[0]](arguments[1]); }", callback, captcha_token)
            logger.info(f"Callback '{callback}' acionado com o token.")
        else:
            logger.info("Nenhum callback definido para o reCAPTCHA.")
    except Exception as e:
        logger.error("Erro ao injetar token ou disparar callback: " + str(e))
    return captcha_token

def main():
    try:
        chrome_driver_path = "/home/lfragoso/projetos/dash-burgetXLogComexXComercial/chromedriver"
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument("--remote-debugging-port=0")

        # Log auxiliar para confirmar os argumentos ativos (opcional)
        for arg in chrome_options.arguments:
            logger.debug(f"[DEBUG] Chrome argument ativo: {arg}")

        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        wait = WebDriverWait(driver, 30)

        try:
            driver.get("https://plataforma.logcomex.io/signIn/")
            logger.info("Acessando o site de login.")
            WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")
            email_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-23"]')))
            password_field = driver.find_element(By.XPATH, '//*[@id="input-27"]')
            logger.info("Campos de email e senha encontrados.")
            email = "Apoiocomercial.rj@ictsirio.com"
            password = "Apoio*321"
            slow_typing(email_field, email)
            slow_typing(password_field, password)
            logger.info("Campos preenchidos, iniciando resolução do reCAPTCHA.")
            
            api_key = 'bb697646215cc9c54062c09f063e093f'
            try:
                captcha_token = solve_recaptcha(driver, api_key)
                logger.info("Token do captcha obtido: " + captcha_token)
            except Exception as e:
                logger.warning("Não foi possível resolver o captcha automaticamente: " + str(e))
            time.sleep(5)
            
            # 1) Clica no card "Shipment Intel" (botão "Acessar")
            acessar_shipment_intel(driver, wait)
            
            # 2) Navega para "Importação Marítima" dentro do Shipment Intel e aguarda o carregamento da seção
            navegar_para_secao(driver, wait)
            WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")
            ajustar_zoom(driver)
            time.sleep(1)
            
            # 3) Aplicação dos filtros para importação
            filter_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="button_side_menu_filter"]')))
            filter_button.click()
            logger.info("Botão de filtros clicado.")
            time.sleep(2)
            
            retry_count = 3
            for attempt in range(retry_count):
                try:
                    master_filter = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="box_embarque_62c3d669-c1ed-499e-bcd5-32ff1108a814"]/div/div[2]/div/div/div[4]/div/div[1]/label')
                    ))
                    ActionChains(driver).move_to_element(master_filter).click().perform()
                    logger.info("Filtro Master selecionado.")
                    break
                except Exception as e:
                    logger.warning(f"Tentativa {attempt + 1}/{retry_count}: Erro ao clicar no filtro Master.")
                    time.sleep(2)
            else:
                raise Exception("Falha ao interagir com o filtro Master após múltiplas tentativas.")
            
            for attempt in range(retry_count):
                try:
                    porto_destino = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="title_cod_nome_porto_destino_62c3d669-c1ed-499e-bcd5-32ff1108a814"]')
                    ))
                    ActionChains(driver).move_to_element(porto_destino).click().perform()
                    logger.info("Campo Porto de Destino aberto.")
                    time.sleep(3)
                    break
                except Exception as e:
                    logger.warning(f"Tentativa {attempt + 1}/{retry_count}: Porto de Destino não interativo.")
                    time.sleep(2)
            else:
                raise Exception("Falha ao interagir com o campo Porto de Destino após múltiplas tentativas.")
            
            porto_destino_box = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="box_cod_nome_porto_destino_62c3d669-c1ed-499e-bcd5-32ff1108a814"]/div/div[2]/div/div/div/div/div/div[1]/div[1]'))
            )
            driver.execute_script("arguments[0].click();", porto_destino_box)
            logger.info("Campo de digitação para Porto de Destino clicado.")
            time.sleep(1)
            ActionChains(driver).send_keys("BRRIO RIO DE JANEIRO").send_keys(Keys.ENTER).perform()
            time.sleep(1)
            ActionChains(driver).send_keys("BRIGI ITAGUAI").send_keys(Keys.ENTER).perform()
            logger.info("Portos de destino inseridos.")
            
            # 4) Seleção de período via calendário utilizando o mês atual inteiro
            logger.info("Abrindo o calendário para selecionar período.")
            calendario = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="input-118"]')))
            calendario.click()
            time.sleep(1)
            calendario.send_keys(Keys.CONTROL + "a")
            calendario.send_keys(Keys.DELETE)
            time.sleep(0.5)
            inicio_data, fim_data = get_date_range_current_month()
            calendario.send_keys(f"{inicio_data} {fim_data}")
            time.sleep(1)
            calendario.send_keys(Keys.ENTER)
            logger.info(f"Período {inicio_data} até {fim_data} inserido e aplicado.")
            time.sleep(2)
            
            # 5) Clique nos botões (Ok, Filtrar, Detalhes)
            buttons = {
                'Ok': '//*[@id="LSSkeletonCookie__agreeButton"]',
                'Filtrar': '//*[@id="button_filter"]',
                'Detalhes': '//*[@id="button_side_menu_details"]'
            }
            for button_name, xpath in buttons.items():
                btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                btn.click()
                logger.info(f"Botão {button_name} clicado.")
                time.sleep(1)
            
            driver.execute_script("document.body.style.zoom='67%'")
            logger.info("Zoom ajustado para 67%.")
            time.sleep(1)
            
            # 6) Extração dos dados da tabela
            table = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pdf-export-container"]/div[2]')))
            logger.info("Tabela encontrada. Iniciando extração de dados.")
            time.sleep(5)
            
            # Verificação de mensagem de ausência de registros
            no_records_xpath = '//*[@id="pdf-export-container"]/div[2]/div/div[1]/div/div[2]/div[2]'
            time.sleep(2)
            if driver.find_elements(By.XPATH, no_records_xpath):
                mensagem_no_result = driver.find_element(By.XPATH, no_records_xpath).text.strip()
                if "Nenhum resultado foi encontrado na sua pesquisa" in mensagem_no_result:
                    logger.warning("Nenhum resultado foi encontrado na sua pesquisa. Gerando planilha informando ausência de dados.")
                    df_vazio = pd.DataFrame({"Mensagem": [f"Sem dados para o dia informado: {fim_data}"]})
                    nome_arquivo = 'Importação.xlsx'
                    salvar_arquivo_excel(df_vazio, nome_arquivo)
                    logger.info("Planilha gerada informando ausência de dados. Encerrando extração.")
                    driver.quit()
                    return
            else:
                logger.info("Mensagem de 'Nenhum resultado' não encontrada. Prosseguindo com a extração.")
            
            header_cells = wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//*[@id='pdf-export-container']/div[2]//div[contains(@class, 'ag-header-cell-text')]")
            ))
            header_texts = [header.get_attribute('textContent').strip() for header in header_cells if header.get_attribute('textContent').strip()]
            logger.info(f"Cabeçalhos capturados: {header_texts}")
            if not header_texts:
                raise Exception("Nenhum cabeçalho encontrado na tabela")
            
            all_data = []
            total_pages = get_total_pages(driver)
            current_page = 1
            while current_page <= total_pages:
                logger.info(f"Processando página {current_page}")
                time.sleep(2)
                try:
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ag-center-cols-container")))
                    pagination_text = driver.find_element(By.CSS_SELECTOR, "span.ag-paging-description").text
                    numeros = re.findall(r'\d+', pagination_text)
                    if len(numeros) >= 2:
                        pagina_atual, total_pages = int(numeros[0]), int(numeros[1])
                        logger.info(f"Página {pagina_atual} de {total_pages}")
                    else:
                        logger.error("Erro ao extrair números de paginação")
                        break
                    
                    scroll_table_fully(driver)
                    time.sleep(1)
                    page_data = extract_table_data(driver, wait, header_texts)
                    max_attempts = 3
                    attempt = 0
                    while len(page_data) < 30 and attempt < max_attempts:
                        logger.warning(f"Tentativa {attempt + 1}: Menos de 30 registros encontrados. Realizando nova rolagem.")
                        scroll_table_fully(driver)
                        time.sleep(1)
                        page_data = extract_table_data(driver, wait, header_texts)
                        attempt += 1
                    if page_data:
                        all_data.extend(page_data)
                        logger.info(f"Extraídos {len(page_data)} registros da página {current_page}")
                    else:
                        logger.warning(f"Nenhum dado encontrado na página {current_page}")
                        driver.save_screenshot(f"pagina_vazia_{current_page}.png")
                    
                    if not navigate_to_next_page(driver, wait, current_page):
                        logger.info("Última página alcançada ou falha na navegação.")
                        break
                    current_page += 1
                except Exception as e:
                    logger.error(f"Erro ao processar página {current_page}: {e}")
                    driver.save_screenshot(f"erro_pagina_{current_page}.png")
                    break
            
            # 7) Processa e salva os dados, caso existam
            if all_data:
                try:
                    df = process_data_frame(all_data, header_texts)
                    nome_arquivo = 'Importação.xlsx'
                    salvar_arquivo_excel(df, nome_arquivo)
                except Exception as e:
                    logger.error(f"Erro ao processar e salvar dados: {e}")
                    raise
            else:
                logger.warning("Nenhum dado coletado para importação.")
        except Exception as e:
            logger.error(f"Erro durante a execução: {e}")
            driver.save_screenshot("erro_execucao.png")
            raise
        finally:
            logger.info("Finalizando o script.")
            driver.quit()
    except Exception as e:
        logger.error(f"Erro crítico: {e}")
        raise

if __name__ == "__main__":
    main()
