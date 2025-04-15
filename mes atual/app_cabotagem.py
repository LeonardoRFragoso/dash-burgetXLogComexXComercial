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

class CaptchaError(Exception):
    """Erro ao resolver captcha"""
    pass

# ----------------- Funções de Data -----------------
# Funções originais (mantidas para referência ou uso posterior)
def get_date_range_7_days():
    hoje = datetime.now()
    fim = hoje - timedelta(days=1)  # Dia anterior
    inicio = hoje - timedelta(days=7)  # Período de 7 dias
    return inicio.strftime("%d-%m-%Y"), fim.strftime("%d-%m-%Y")

def get_date_range():
    hoje = datetime.now()
    fim = hoje - timedelta(days=1)  # Dia anterior
    inicio = hoje - timedelta(days=30)  # Período de 30 dias
    return inicio.strftime("%d-%m-%Y"), fim.strftime("%d-%m-%Y")

def get_last_30_days_dates():
    hoje = datetime.now()
    dia_anterior = hoje - timedelta(days=1)
    inicio_periodo = dia_anterior - timedelta(days=29)  # Total de 30 dias
    return inicio_periodo.strftime("%d-%m-%Y"), dia_anterior.strftime("%d-%m-%Y")

# NOVA FUNÇÃO: Retorna o intervalo de datas desde o primeiro dia do mês atual até o dia anterior.
def get_date_range_current_month():
    hoje = datetime.now()
    fim = hoje - timedelta(days=1)  # Dia anterior
    inicio = datetime(hoje.year, hoje.month, 1)  # Primeiro dia do mês atual
    return inicio.strftime("%d-%m-%Y"), fim.strftime("%d-%m-%Y")

# ----------------- Funções Auxiliares -----------------
def slow_typing(element, text, delay=0.1):
    for char in text:
        element.send_keys(char)
        time.sleep(delay)

# ----------------- Funções de Extração da Tabela -----------------
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
            logging.info(f"{num_rows} linhas atualmente visíveis.")
            if num_rows >= 30:
                logging.info("30 linhas capturadas, rolagem completa.")
                break
            if num_rows == last_row_count:
                actions.scroll_by_amount(0, 50).perform()
            else:
                actions.move_to_element(rows[-1]).perform()
            last_row_count = num_rows
            time.sleep(0.3)
            attempt += 1
        if attempt == max_attempts:
            logging.warning("Máximo de tentativas de rolagem atingido. Linhas ainda incompletas.")
    except Exception as e:
        logging.warning(f"Erro durante a rolagem: {str(e)}")

def wait_for_table_load(driver, timeout=15):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ag-center-cols-container"))
    )
    logging.info("Tabela completamente carregada.")

def extract_table_data(driver, wait, header_texts, max_retries=3):
    for attempt in range(max_retries):
        try:
            logging.info("Tentando extrair dados da tabela usando JavaScript...")
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
                logging.info(f"{len(rows)} linhas capturadas.")
                page_data = []
                for row in rows:
                    while len(row) < len(header_texts):
                        row.append('')
                    if len(row) == len(header_texts):
                        page_data.append(row)
                    else:
                        logging.warning(f"Linha inconsistente, preenchida com valores vazios: {row}")
                return page_data
        except Exception as e:
            driver.save_screenshot(f"erro_extracao_{attempt + 1}.png")
            logging.warning(f"Tentativa {attempt + 1} falhou: {str(e)}")
            time.sleep(2)
    logging.error("Falha após múltiplas tentativas de extração.")
    return []

def get_total_pages(driver):
    try:
        pagination_xpath = "//span[@class='ag-paging-description']"
        pagination_text = driver.find_element(By.XPATH, pagination_xpath).text
        total_pages = int(re.search(r"de (\d+)", pagination_text).group(1))
        logging.info(f"Total de páginas detectadas: {total_pages}")
        return total_pages
    except Exception as e:
        logging.warning(f"Falha ao obter o total de páginas: {str(e)}")
        return 1

def navigate_to_next_page(driver, wait, current_page):
    try:
        next_button_container = driver.find_element(By.XPATH, "//*[@id='ag-42']")
        next_button = next_button_container.find_element(By.CLASS_NAME, "ag-icon-next")
        if next_button.is_displayed() and next_button.is_enabled():
            next_button.click()
            logging.info(f"Processando página {current_page + 1}")
            time.sleep(2)
            return True
        else:
            logging.info("Última página alcançada.")
            return False
    except Exception as e:
        logging.error(f"Erro ao tentar clicar na próxima página: {e}")
        return False

def process_data_frame(all_data, header_texts):
    try:
        df = pd.DataFrame(all_data, columns=header_texts)
        data_atual = datetime.now().strftime('%d/%m/%Y')
        df['DATA CONSULTA'] = data_atual
        df['DATA EXTRACAO'] = data_atual
        df['DATA CONSULTA'] = df['DATA CONSULTA'].astype(str)
        df['DATA EXTRACAO'] = df['DATA EXTRACAO'].astype(str)
        df = df.drop_duplicates()
        df = df.dropna(how='all')
        logging.info(f"DataFrame após limpeza: {df.shape}")
        df = corrigir_tipos_dados(df)
        return df
    except Exception as e:
        logging.error(f"Erro no processamento do DataFrame: {str(e)}")
        raise

def corrigir_tipos_dados(df):
    try:
        logging.info(f"Iniciando correção de tipos de dados. Shape inicial: {df.shape}")
        campos_numericos = [
            'ID', 'TRANSIT-TIME', 'TEMPO PORTO', 'QTDE CONSIGNATÁRIO FINAL',
            'QUANTIDADE VEICULOS', 'QTDE CONTAINER', 'TEUS', 'C20', 'C40',
            'VOLUMES', 'PESO BRUTO'
        ]
        for campo in campos_numericos:
            if campo in df.columns:
                try:
                    df[campo] = pd.to_numeric(
                        df[campo].str.replace(',', '.').str.replace('R$', '').str.strip(),
                        errors='coerce'
                    )
                    df[campo] = df[campo].fillna(0)
                    logging.info(f"Campo {campo} convertido para numérico")
                except Exception as e:
                    logging.warning(f"Erro ao converter campo {campo}: {str(e)}")
        campos_data = ['ETS', 'ETA', 'SAÍDA PORTO']
        for campo in campos_data:
            if campo in df.columns:
                try:
                    df[campo] = pd.to_datetime(df[campo], format='%d/%m/%Y', errors='coerce')
                    df[campo] = df[campo].dt.strftime('%d/%m/%Y')
                    logging.info(f"Campo {campo} convertido para data")
                except Exception as e:
                    logging.warning(f"Erro ao converter campo {campo}: {str(e)}")
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
            df['ANO/MÊS'] = df['ANO/MÊS'].fillna(ano_mes_atual)
            df['ANO/MÊS'] = df['ANO/MÊS'].astype(int)
            logging.info("Coluna ANO/MÊS processada e convertida para inteiro")
        if 'DATA CONSULTA' in df.columns:
            df['DATA CONSULTA'] = df['DATA CONSULTA'].astype(str)
            logging.info("Coluna DATA CONSULTA convertida para string")
        logging.info(f"Correção de tipos concluída. Shape final: {df.shape}")
        return df
    except Exception as e:
        logging.error(f"Erro no processamento do DataFrame: {str(e)}")
        logging.error(f"Colunas do DataFrame: {df.columns.tolist()}")
        raise

def retry_action(action, error_message="Erro ao interagir com elemento", max_retries=3):
    for attempt in range(max_retries):
        try:
            return action()
        except Exception as e:
            logging.warning(f"{error_message}. Tentativa {attempt + 1}/{max_retries}.")
            time.sleep(2)
    raise Exception(f"{error_message} após {max_retries} tentativas.")

# ----------------- Funções de reCAPTCHA -----------------
def solve_recaptcha(driver, api_key):
    try:
        recaptcha_button = driver.find_element(By.ID, "signIn__container__card__form__signInButton")
        site_key = recaptcha_button.get_attribute("data-sitekey")
        current_url = driver.current_url
        logging.info(f"Site key encontrada: {site_key}")
    except Exception as e:
        logging.error(f"Não foi possível localizar o botão com recaptcha: {e}")
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
    logging.info(f"Captcha enviado para 2Captcha, ID: {captcha_id}")
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
            logging.info("Captcha resolvido com sucesso.")
            break
        else:
            logging.info(f"Aguardando solução do captcha... tentativa {attempt + 1}")
            time.sleep(5)
    if not captcha_token:
        raise Exception("Tempo excedido para resolução do captcha.")
    try:
        driver.execute_script("document.getElementById('g-recaptcha-response').style.display='block';")
        driver.execute_script("document.getElementById('g-recaptcha-response').value = arguments[0];", captcha_token)
        driver.execute_script(
            "var event = new Event('input', { bubbles: true });" +
            "document.getElementById('g-recaptcha-response').dispatchEvent(event);",
            captcha_token
        )
        logging.info("Token injetado no g-recaptcha-response.")
        callback = driver.execute_script(
            "return document.getElementById('signIn__container__card__form__signInButton').getAttribute('data-callback');"
        )
        if callback:
            driver.execute_script("if(window[arguments[0]]) { window[arguments[0]](arguments[1]); }", callback, captcha_token)
            logging.info(f"Callback '{callback}' acionado com o token.")
        else:
            logging.info("Nenhum callback definido para o reCAPTCHA.")
    except Exception as e:
        logging.error("Erro ao injetar token ou disparar callback: " + str(e))
    return captcha_token

# ----------------- Funções de Filtros (lógica original de Cabotagem) -----------------
def aplicar_filtro_data(driver, wait):
    try:
        logging.info("Abrindo o calendário para selecionar período.")
        # Localiza o campo de data
        calendario = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="input-118"]')))
        calendario.click()
        time.sleep(1)
        # Seleciona todo o conteúdo e apaga
        calendario.send_keys(Keys.CONTROL + "a")
        calendario.send_keys(Keys.DELETE)
        time.sleep(0.5)
        # Obtém as datas: do primeiro dia do mês atual até o dia anterior
        inicio_data, fim_data = get_date_range_current_month()
        calendario.send_keys(f"{inicio_data} {fim_data}")
        time.sleep(1)
        # Clica no botão de aplicar filtro
        aplicar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[2]/footer/button/span')))
        aplicar_btn.click()
        logging.info(f"Período {inicio_data} até {fim_data} inserido e aplicado.")
        time.sleep(2)
    except Exception as e:
        logging.error(f"Erro ao aplicar filtro de data: {str(e)}")
        driver.save_screenshot("erro_filtro_data.png")
        raise

def aplicar_filtros(driver, wait):
    try:
        logging.info("Iniciando aplicação de filtros adicionais.")
        
        # Clicar no botão de filtros
        filter_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="button_side_menu_filter"]')))
        filter_button.click()
        logging.info("Botão de filtros clicado.")
        time.sleep(2)
        
        # Selecionar tipos de carga
        tipo_carga = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="title_tp_carga_1d23bd0c-f7f1-4fa7-ba05-4042ec22ea7d"]')))
        tipo_carga.click()
        time.sleep(1)
        tipos_carga_xpath = {
            'break bulk': '//*[@id="box_tp_carga_1d23bd0c-f7f1-4fa7-ba05-4042ec22ea7d"]/div/div[2]/div/div/div[1]/div/div[1]/label',
            'container': '//*[@id="box_tp_carga_1d23bd0c-f7f1-4fa7-ba05-4042ec22ea7d"]/div/div[2]/div/div/div[2]/div/div[1]/label',
            'solta': '//*[@id="box_tp_carga_1d23bd0c-f7f1-4fa7-ba05-4042ec22ea7d"]/div/div[2]/div/div/div[6]/div/div[1]/label',
            'sólidos': '//*[@id="box_tp_carga_1d23bd0c-f7f1-4fa7-ba05-4042ec22ea7d"]/div/div[2]/div/div/div[5]/div/div[1]/label'
        }
        for tipo, xpath in tipos_carga_xpath.items():
            try:
                elemento = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                elemento.click()
                logging.info(f"Tipo de carga '{tipo}' selecionado.")
                time.sleep(0.5)
            except Exception as e:
                logging.warning(f"Não foi possível selecionar o tipo de carga '{tipo}': {str(e)}")
        
        # Aplicar filtro de Estado do Destinatário
        try:
            estado_filtro = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="title_uf_consignatario_1d23bd0c-f7f1-4fa7-ba05-4042ec22ea7d"]')))
            estado_filtro.click()
            logging.info("Filtro de Estado do Destinatário aberto.")
            time.sleep(1)
        except Exception as e:
            logging.error("Não foi possível clicar no filtro de Estado do Destinatário.")
            raise

        estados_destinatario = ["RJ", "MG", "SP", "ES", "SC", "PR"]
        campo_estado_xpath = '//*[@id="box_uf_consignatario_1d23bd0c-f7f1-4fa7-ba05-4042ec22ea7d"]/div/div[2]/div/div/div/div/div/div[1]/div[1]'
        digitavel_campo_estado_element = wait.until(EC.element_to_be_clickable((By.XPATH, campo_estado_xpath)))
        digitavel_campo_estado_element.click()
        time.sleep(1)
        
        for estado in estados_destinatario:
            ActionChains(driver).send_keys(estado).send_keys(Keys.ENTER).perform()
            logging.info(f"Estado {estado} inserido.")
            time.sleep(0.5)

        # Aplicar o filtro de data após os demais filtros
        aplicar_filtro_data(driver, wait)
        
        logging.info("Filtros aplicados com sucesso.")
    except Exception as e:
        logging.error(f"Erro ao aplicar filtros: {str(e)}")
        driver.save_screenshot("erro_filtros.png")
        raise

# ----------------- Função de Navegação para Cabotagem -----------------
def navegar_para_secao_cabotagem(driver, wait):
    try:
        menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LSTopBreadcrumbMenuActivator"]/div[1]')))
        menu_button.click()
        logging.info("Menu Shipment Intel clicado.")
        time.sleep(2)
        cabotagem = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[2]/div/div/div/div[1]/div[2]/ul/li[1]')))
        cabotagem.click()
        logging.info("Cabotagem selecionada.")
        time.sleep(2)
        brasil = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[2]/div/div/div/div[2]/div[2]/ul/li')))
        brasil.click()
        logging.info("Brasil selecionado.")
        time.sleep(2)
    except Exception as e:
        logging.error(f"Erro ao navegar para cabotagem: {str(e)}")
        driver.save_screenshot("erro_navegacao_cabotagem.png")
        raise

# ----------------- Funções de Resumo e Salvamento -----------------
def adicionar_aba_resumo(df, file_name):
    try:
        resumo_df = df.iloc[:, [3, 10, 17, 30]].copy()
        resumo_df.insert(0, 'DATA CONSULTA', datetime.now().strftime('%d/%m/%Y'))
        resumo_df.columns = ['DATA CONSULTA', 'DATA EMBARQUE', 'PORTO EMBARQUE', 'NOME EXPORTADOR', 'QTDE CONTAINER']
        
        def convert_container_qty(value):
            try:
                value_str = str(value).replace(',', '.')
                return float(value_str)
            except:
                return 0.0

        resumo_df['QTDE CONTAINER'] = resumo_df['QTDE CONTAINER'].apply(convert_container_qty)
        
        # Converter a coluna DATA EMBARQUE para datetime, tratando valores inválidos
        resumo_df['DATA EMBARQUE'] = pd.to_datetime(resumo_df['DATA EMBARQUE'], errors='coerce')
        
        # Remover linhas onde a data é inválida (NaT)
        linhas_removidas = resumo_df['DATA EMBARQUE'].isna().sum()
        if linhas_removidas > 0:
            logging.warning(f"Removidas {linhas_removidas} linhas com data de embarque inválida")
        
        resumo_df = resumo_df.dropna(subset=['DATA EMBARQUE'])
        
        # Agora podemos formatar as datas válidas com segurança
        resumo_df['DATA EMBARQUE'] = resumo_df['DATA EMBARQUE'].dt.strftime('%d/%m/%Y')
        
        logging.info(f"\nAmostra antes do agrupamento (após limpeza de datas): {len(resumo_df)} registros")
        logging.info(resumo_df.head().to_string())
        
        resumo_agrupado = resumo_df.groupby(['NOME EXPORTADOR', 'DATA CONSULTA', 'DATA EMBARQUE', 'PORTO EMBARQUE'], as_index=False).agg({'QTDE CONTAINER': 'sum'})
        resumo_agrupado = resumo_agrupado.sort_values(['NOME EXPORTADOR', 'DATA CONSULTA', 'DATA EMBARQUE', 'PORTO EMBARQUE']).reset_index(drop=True)
        
        logging.info("\nAmostra após agrupamento:")
        logging.info(resumo_agrupado.head().to_string())
        
        resumo_agrupado['QTDE CONTAINER'] = resumo_agrupado['QTDE CONTAINER'].apply(lambda x: f"{x:.2f}".replace('.', ','))
        
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
            resumo_agrupado.to_excel(writer, sheet_name='Resumo', index=False)
            logging.info(f"Aba de resumo atualizada com sucesso: {file_name}")
        
        return resumo_agrupado
    except Exception as e:
        logging.error(f"Erro ao adicionar aba de resumo: {str(e)}")
        logging.info("Continuando a execução mesmo sem adicionar a aba de resumo.")
        return None

def ajustar_zoom(driver, zoom_level=80):
    try:
        driver.execute_script(f"document.body.style.zoom = '{zoom_level}%'")
        logging.info(f"Zoom ajustado para {zoom_level}%")
    except Exception as e:
        logging.warning(f"Não foi possível ajustar o zoom: {e}")

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
            logging.info(f"Backup criado: {nome_backup}")
        except Exception as e:
            logging.warning(f"Não foi possível criar backup: {e}")
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
                logging.warning(f"Arquivo {nome_arquivo} está em uso. Tentativa {tentativa}/{max_tentativas}")
                time.sleep(5)
                tentativa += 1
                continue
            with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Logcomex', index=False)
                logging.info(f"Aba Logcomex exportada com sucesso: {len(df)} registros")
            if os.path.exists(nome_arquivo) and os.path.getsize(nome_arquivo) > 0:
                logging.info(f"Arquivo {nome_arquivo} salvo com sucesso")
                try:
                    adicionar_aba_resumo(df, nome_arquivo)
                    logging.info("Aba de resumo adicionada com sucesso")
                except Exception as e:
                    logging.error(f"Erro ao adicionar aba de resumo: {e}")
                if os.path.exists(nome_backup):
                    os.remove(nome_backup)
                    logging.info("Arquivo de backup removido")
                break
            else:
                raise Exception("Arquivo não foi salvo corretamente")
        except Exception as e:
            logging.error(f"Erro ao salvar arquivo (tentativa {tentativa}): {e}")
            if tentativa == max_tentativas:
                if os.path.exists(nome_backup):
                    if os.path.exists(nome_arquivo):
                        os.remove(nome_arquivo)
                    os.rename(nome_backup, nome_arquivo)
                    logging.info("Backup restaurado após falha")
                raise
            tentativa += 1
            time.sleep(5)

# ----------------- Função Main -----------------
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
            logging.info("Acessando o site de login.")
            WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")

            email_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-23"]')))
            password_field = driver.find_element(By.XPATH, '//*[@id="input-27"]')
            logging.info("Campos de email e senha encontrados.")

            email = "Apoiocomercial.rj@ictsirio.com"
            password = "Apoio*321"
            
            slow_typing(email_field, email)
            slow_typing(password_field, password)
            
            try:
                captcha_token = solve_recaptcha(driver, "bb697646215cc9c54062c09f063e093f")
                logging.info("Token do captcha obtido: " + captcha_token)
            except Exception as e:
                logging.warning("Não foi possível resolver o captcha automaticamente: " + str(e))
            time.sleep(5)  # Aguarda para que o token seja processado pelo site
            
            # Agora, clica no botão de login (ENTRAR) apenas após o token estar injetado
            button_acessar = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="app"]/div/div/main/div/div/div[1]/div[1]/div[3]/div/div/div/div[1]/div[2]/button')
            ))
            button_acessar.click()
            logging.info("Botão Acessar clicado.")
            time.sleep(2)

            # Navegar para seção de cabotagem
            navegar_para_secao_cabotagem(driver, wait)

            # Esperar carregamento completo
            WebDriverWait(driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )

            # Ajusta o zoom da página para 80% antes de aplicar os filtros
            ajustar_zoom(driver)
            time.sleep(1)  # Pequena pausa para garantir que o zoom foi aplicado

            # Aplicar filtros de carga, estado e porto de descarga
            aplicar_filtros(driver, wait)

            # Cliques nos botões necessários
            buttons = {
                'Ok': '//*[@id="LSSkeletonCookie__agreeButton"]',
                'Filtrar': '//*[@id="button_filter"]',
                'Detalhes': '//*[@id="button_side_menu_details"]'
            }

            for button_name, xpath in buttons.items():
                try:
                    button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    button.click()
                    logging.info(f"Botão {button_name} clicado.")
                    time.sleep(2)  # Aumentado o tempo de espera entre cliques
                    
                    # Adicionar espera especial após clicar em Filtrar e Detalhes
                    if button_name in ['Filtrar', 'Detalhes']:
                        logging.info(f"Aguardando carregamento completo após clicar em {button_name}...")
                        time.sleep(5)  # Pausa inicial para dar tempo ao JavaScript carregar
                        
                        # Esperar até que o loading spinner desapareça (se existir)
                        try:
                            WebDriverWait(driver, 10).until_not(
                                EC.presence_of_element_located((By.CLASS_NAME, "loading-spinner"))
                            )
                        except:
                            logging.info("Loading spinner não encontrado ou já desapareceu")
                        
                        # Esperar até que a tabela esteja presente e visível, especialmente após clicar em Detalhes
                        if button_name == 'Detalhes':
                            try:
                                WebDriverWait(driver, 20).until(
                                    EC.visibility_of_element_located((By.CLASS_NAME, "ag-center-cols-container"))
                                )
                                logging.info("Tabela principal carregada e visível")
                            except Exception as e:
                                logging.warning(f"Aviso: Tempo limite excedido aguardando tabela: {e}")
                        
                        # Esperar mais alguns segundos para garantir que todos os dados foram carregados
                        time.sleep(3)
                except Exception as e:
                    logging.warning(f"Não foi possível clicar no botão {button_name}: {e}")
                    if button_name != 'Ok':
                        raise

            # Verificação de mensagem de ausência de registros
            no_records_xpath = '//*[@id="pdf-export-container"]/div[2]/div/div[1]/div/div[2]/div[2]'
            time.sleep(2)  # Aguarda um instante para a mensagem ser exibida, se houver
            if driver.find_elements(By.XPATH, no_records_xpath):
                mensagem_no_result = driver.find_element(By.XPATH, no_records_xpath).text.strip()
                if "Nenhum resultado foi encontrado na sua pesquisa" in mensagem_no_result:
                    logging.warning("Mensagem de ausência de registros encontrada: " + mensagem_no_result)
                    # Utiliza a nova função para capturar a data final utilizada no filtro
                    _, fim_data = get_date_range_current_month()
                    df_vazio = pd.DataFrame({"Mensagem": [f"Sem dados para o dia informado: {fim_data}"]})
                    nome_arquivo = 'Cabotagem.xlsx'
                    salvar_arquivo_excel(df_vazio, nome_arquivo)
                    logging.info("Planilha gerada informando ausência de dados. Encerrando extração.")
                    driver.quit()
                    return
            else:
                logging.info("Mensagem de 'Nenhum resultado' não encontrada. Prosseguindo com a extração.")
            
            # Iniciar extração dos dados da tabela
            table = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pdf-export-container"]/div[2]')))
            logging.info("Tabela encontrada. Iniciando extração de dados.")
            time.sleep(5)
            
            header_cells = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//*[@id='pdf-export-container']/div[2]//div[contains(@class, 'ag-header-cell-text')]")))
            header_texts = [header.get_attribute('textContent').strip() for header in header_cells if header.get_attribute('textContent').strip()]
            logging.info(f"Cabeçalhos capturados: {header_texts}")
            
            if not header_texts:
                raise Exception("Nenhum cabeçalho encontrado na tabela")
            
            all_data = []
            total_pages = get_total_pages(driver)
            current_page = 1
            
            while current_page <= total_pages:
                logging.info(f"Processando página {current_page}")
                time.sleep(2)
                try:
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ag-center-cols-container")))
                    pagination_text = driver.find_element(By.CSS_SELECTOR, "span.ag-paging-description").text
                    numeros = re.findall(r'\d+', pagination_text)
                    if len(numeros) >= 2:
                        pagina_atual, total_pages = int(numeros[0]), int(numeros[1])
                        logging.info(f"Página {pagina_atual} de {total_pages}")
                    else:
                        logging.error("Erro ao extrair números de paginação")
                        break
                    
                    scroll_table_fully(driver)
                    time.sleep(1)
                    page_data = extract_table_data(driver, wait, header_texts)
                    max_attempts = 3
                    attempt = 0
                    while len(page_data) < 30 and attempt < max_attempts:
                        logging.warning(f"Tentativa {attempt + 1}: Menos de 30 registros encontrados. Realizando nova rolagem.")
                        scroll_table_fully(driver)
                        time.sleep(1)
                        page_data = extract_table_data(driver, wait, header_texts)
                        attempt += 1
                    if page_data:
                        all_data.extend(page_data)
                        logging.info(f"Extraídos {len(page_data)} registros da página {current_page}")
                    else:
                        logging.warning(f"Nenhum dado encontrado na página {current_page}")
                        driver.save_screenshot(f"pagina_vazia_{current_page}.png")
                    
                    if not navigate_to_next_page(driver, wait, current_page):
                        logging.info("Última página alcançada ou falha na navegação.")
                        break
                    current_page += 1
                except Exception as e:
                    logging.error(f"Erro ao processar página {current_page}: {e}")
                    driver.save_screenshot(f"erro_pagina_{current_page}.png")
                    break
            
            if all_data:
                try:
                    df = process_data_frame(all_data, header_texts)
                    nome_arquivo = 'Cabotagem.xlsx'
                    salvar_arquivo_excel(df, nome_arquivo)
                except Exception as e:
                    logging.error(f"Erro ao processar e salvar dados: {e}")
                    raise
            else:
                logging.warning("Nenhum dado coletado para exportação.")
                
        except Exception as e:
            logging.error(f"Erro durante a execução: {e}")
            driver.save_screenshot("erro_execucao.png")
            raise
            
        finally:
            driver.quit()
            
    except Exception as e:
        logging.error(f"Erro crítico: {e}")
        raise

if __name__ == "__main__":
    main()
