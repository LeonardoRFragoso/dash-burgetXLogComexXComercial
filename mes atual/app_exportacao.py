import time
import requests
import tempfile
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

# ----------------- Funções de Data -----------------
def get_date_range_7_days():
    hoje = datetime.now()
    fim = hoje - timedelta(days=1)  # D-1 (dia anterior ao atual)
    inicio = hoje - timedelta(days=7)  # Período de 7 dias
    return inicio.strftime("%d-%m-%Y"), fim.strftime("%d-%m-%Y")

def get_date_range():
    hoje = datetime.now()
    fim = hoje - timedelta(days=1)  # D-1 (dia anterior ao atual)
    inicio = hoje - timedelta(days=30)  # Período de 30 dias
    return inicio.strftime("%d-%m-%Y"), fim.strftime("%d-%m-%Y")

# NOVA FUNÇÃO: Retorna intervalo de datas a partir do primeiro dia do mês atual até o dia anterior.
def get_date_range_current_month():
    hoje = datetime.now()
    fim = hoje - timedelta(days=1)
    inicio = datetime(hoje.year, hoje.month, 1)
    return inicio.strftime("%d-%m-%Y"), fim.strftime("%d-%m-%Y")

# ----------------- Funções de rolagem e extração da tabela -----------------
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

def capture_all_rows(driver, wait, expected_rows=30):
    all_rows = []
    wait_for_table_load(driver)
    total_rows = driver.execute_script("return document.querySelectorAll('#pdf-export-container .ag-center-cols-container .ag-row').length;")
    while len(all_rows) < expected_rows:
        rows = driver.execute_script("""
        let rows = document.querySelectorAll('#pdf-export-container .ag-center-cols-container .ag-row');
        let rowData = [];
        rows.forEach(row => {
            let cells = row.querySelectorAll('.ag-cell');
            let data = Array.from(cells).map(cell => cell.textContent.trim());
            rowData.push(data);
        });
        return rowData;
        """)
        all_rows.extend(rows)
        scroll_table_fully(driver)
        if len(all_rows) >= total_rows:
            break
    if len(all_rows) > expected_rows:
        all_rows = all_rows[:expected_rows]
    logging.info(f"{len(all_rows)} linhas capturadas após rolagem completa.")
    return all_rows

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
        df['DATA_EXTRACAO'] = data_atual
        df['DATA CONSULTA'] = df['DATA CONSULTA'].astype(str)
        df['DATA_EXTRACAO'] = df['DATA_EXTRACAO'].astype(str)
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

def slow_typing(element, text, delay=0.1):
    for char in text:
        element.send_keys(char)
        time.sleep(delay)

# ----------------- Função reCAPTCHA -----------------
def solve_recaptcha(driver, api_key):
    try:
        # Utiliza o botão de login para obter a site key
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
        # Torna o textarea visível para facilitar a injeção do token
        driver.execute_script("document.getElementById('g-recaptcha-response').style.display='block';")
        driver.execute_script("document.getElementById('g-recaptcha-response').value = arguments[0];", captcha_token)
        driver.execute_script(
            "var event = new Event('input', { bubbles: true });"
            "document.getElementById('g-recaptcha-response').dispatchEvent(event);",
            captcha_token
        )
        logging.info("Token injetado no g-recaptcha-response.")
        # Se houver callback definido, executa-o a partir do objeto global window
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

# ----------------- Função de Navegação para Exportação -----------------
def navegar_para_secao(driver, wait):
    try:
        # Clicar no menu Shipment Intel
        menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LSTopBreadcrumbMenuActivator"]/div[1]')))
        menu_button.click()
        logging.info("Menu Shipment Intel clicado.")
        time.sleep(2)
        # Clicar em Exportação Marítima
        exportacao = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[2]/div/div/div/div[1]/div[2]/ul/li[2]')))
        exportacao.click()
        logging.info("Exportação Marítima selecionada.")
        time.sleep(2)
        # Clicar em Brasil
        brasil = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[2]/div/div/div/div[2]/div[2]/ul/li')))
        brasil.click()
        logging.info("Brasil selecionado.")
        time.sleep(2)
    except Exception as e:
        logging.error(f"Erro ao navegar para exportação: {str(e)}")
        driver.save_screenshot("erro_navegacao_simples.png")
        raise

def adicionar_aba_resumo(df, file_name):
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
    resumo_df['DATA EMBARQUE'] = pd.to_datetime(resumo_df['DATA EMBARQUE']).dt.strftime('%d/%m/%Y')
    
    logging.info("\nAmostra antes do agrupamento:")
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

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    try:
        chrome_driver_path = "/home/lfragoso/projetos/dash-burgetXLogComexXComercial/chromedriver"
        chrome_options = Options()
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument("--silent")

        # Adicionar diretório de perfil temporário
        import tempfile
        user_data_dir = tempfile.mkdtemp(prefix="chrome_profile_")
        chrome_options.add_argument(f"--user-data-dir={user_data_dir}")


        
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        wait = WebDriverWait(driver, 20)
        
        try:
            driver.get("https://plataforma.logcomex.io/signIn/")
            logging.info("Acessando o site de login.")
            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            email_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input-23"]')))
            password_field = driver.find_element(By.XPATH, '//*[@id="input-27"]')
            logging.info("Campos de email e senha encontrados.")
            
            email = "Apoiocomercial.rj@ictsirio.com"
            password = "Apoio*321"
            
            slow_typing(email_field, email)
            slow_typing(password_field, password)
            # Removido o envio de RETURN para evitar submissão prematura
            logging.info("Campos preenchidos, iniciando resolução do reCAPTCHA.")
            
            # Resolver o reCAPTCHA via 2Captcha e injetar o token
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
            
            # Navegar para seção de exportação
            navegar_para_secao(driver, wait)
            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            ajustar_zoom(driver)
            time.sleep(1)
            
            # Aplicar filtros
            filter_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="button_side_menu_filter"]')))
            filter_button.click()
            logging.info("Botão de filtros clicado.")
            time.sleep(2)
            
            fechar_tipos_embarque = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="title_embarque_1356f44a-f821-43c4-90e4-dbb04831798a"]')))
            fechar_tipos_embarque.click()
            logging.info("Tipos de Embarque minimizados.")
            
            # Seleção de data dinâmica utilizando o período do primeiro dia do mês atual até o dia anterior
            logging.info("Abrindo o calendário para selecionar período.")
            calendario = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="input-118"]')))
            calendario.click()
            time.sleep(1)
            calendario.send_keys(Keys.CONTROL + "a")
            calendario.send_keys(Keys.DELETE)
            time.sleep(0.5)
            # Aqui é realizada a busca com o novo intervalo de datas
            inicio_data, fim_data = get_date_range_current_month()
            calendario.send_keys(f"{inicio_data} {fim_data}")
            time.sleep(1)
            calendario.send_keys(Keys.ENTER)
            logging.info(f"Período {inicio_data} até {fim_data} inserido e aplicado.")
            time.sleep(2)
            
            retry_action(
                lambda: ActionChains(driver).move_to_element(
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="title_porto_embarque_1356f44a-f821-43c4-90e4-dbb04831798a"]')))
                ).click().perform(),
                "Erro ao abrir campo Porto de Embarque"
            )
            logging.info("Campo Porto de Embarque aberto.")
            
            porto_embarque_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="box_porto_embarque_1356f44a-f821-43c4-90e4-dbb04831798a"]/div/div[2]/div/div/div/div/div/div[1]')))
            driver.execute_script("arguments[0].click();", porto_embarque_box)
            logging.info("Campo de digitação para Porto de Embarque clicado.")
            
            ActionChains(driver).send_keys("RIO DE JANEIRO").send_keys(Keys.ENTER).perform()
            ActionChains(driver).send_keys("ITAGUAI").send_keys(Keys.ENTER).perform()
            logging.info("Portos de Embarque inseridos.")
            
            # Clicar nos botões necessários
            buttons = {
                'Ok': '//*[@id="LSSkeletonCookie__agreeButton"]',
                'Filtrar': '//*[@id="button_filter"]',
                'Detalhes': '//*[@id="button_side_menu_details"]'
            }
            for button_name, xpath in buttons.items():
                retry_action(
                    lambda: wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click(),
                    f"Falha ao clicar no botão {button_name}"
                )
                logging.info(f"Botão {button_name} clicado.")
                time.sleep(1)
            
            # Verificação de mensagem de ausência de registros após clicar em "Detalhes"
            no_records_xpath = '//*[@id="pdf-export-container"]/div[2]/div/div[1]/div/div[2]/div[2]'
            time.sleep(2)  # Aguarda um instante para a mensagem ser exibida, se houver
            if driver.find_elements(By.XPATH, no_records_xpath):
                mensagem = driver.find_element(By.XPATH, no_records_xpath).text
                logging.info(f"Mensagem de ausência de registros encontrada: {mensagem}")
                df_vazio = pd.DataFrame({"Mensagem": [f"Sem dados para o dia informado: {fim_data}"]})
                nome_arquivo = 'Exportação.xlsx'
                salvar_arquivo_excel(df_vazio, nome_arquivo)
                logging.info("Planilha gerada informando ausência de dados. Encerrando extração.")
                driver.quit()
                return
            
            # Aguardar renderização da tabela e extrair dados
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
                    nome_arquivo = 'Exportação.xlsx'
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
