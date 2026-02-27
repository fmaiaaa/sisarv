import re
import time
import random
import unicodedata
from concurrent.futures import ThreadPoolExecutor
import requests
import pandas as pd
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

try:
    from webdriver_manager.chrome import ChromeDriverManager
    USAR_WEBDRIVER_MANAGER = True
except ImportError:
    USAR_WEBDRIVER_MANAGER = False

try:
    from listar_sem_correspondencia import gerar_arquivo_sem_correspondencia
except ImportError:
    gerar_arquivo_sem_correspondencia = None

try:
    from correspondencias_editar import (
        CORRESPONDENCIAS_NOME_POPULAR,
        CORRESPONDENCIAS_NOME_CIENTIFICO,
    )
except ImportError:
    CORRESPONDENCIAS_NOME_POPULAR = {}
    CORRESPONDENCIAS_NOME_CIENTIFICO = {}

base_url = "https://sisarv.rio.gov.br"
# True = utilizar apenas requests (não abre navegador); False = tenta Selenium
USAR_APENAS_REQUESTS = True  # utilizar requests
# True = não preenche o formulário; apenas gera o arquivo com valores sem correspondência no site
NAO_PREENCHER = False

# =============================================================================
# MAPEAMENTO DE PREENCHIMENTO (Coluna DF → Campo no site)
# =============================================================================
# Formato: Campo no site (para preenchimento) = origem do valor.
# Origem = nome da coluna do DF (valor vem da planilha) OU valor fixo (não é coluna).
# Valores fixos (entre chaves na especificação): "", "NÃO", "Espécime não enquadrada...".
# Fácil de editar: altere a coluna ou o valor fixo conforme a necessidade.
#
# Correspondência (Coluna DF - Campo no site):
#   Nº → Nº no Projeto
#   Nome Vulgar → Nome Popular
#   Nome Científico → Nome Científico
#   {VAZIO} → Observação
#   Estado de Conservação → Estado de Conservação
#   Local → Local do Espécime
#   {"Espécime não enquadrada nos casos acima"} → Políticas Municipais
#   {NÃO} → Notabilidade
#   Área Pública → Área Pública
#   Motivação → Motivação
#   Intenção → Intenção
#   H → Altura (m)
#   Copa → Diâmetro da copa(m)
#   DAP 1..5 → DAP 1(cm) .. DAP 5(cm)
#   {NÃO} → Utilidade Pública
# =============================================================================
MAPEAMENTO_PREENCHIMENTO = {
    # Campo no site (preenchimento) : coluna do DF ou valor fixo
    "Nº no Projeto": "Nº",                                     # coluna DF
    "Nome Popular": "Nome Vulgar",                             # coluna DF
    "Nome Científico": "Nome Científico",                      # coluna DF
    "Observação": "",                                          # fixo
    "Estado de Conservação": "Estado de Conservação",          # coluna DF
    "Local do Espécime": "9",                                  # fixo (value da opção "NÃO INFORMADO" no select)
    "Políticas Municipais": "Espécime não enquadrada nos casos acima",  # fixo
    "Notabilidade": "NÃO",                                     # fixo
    "Utilidade Pública": "NÃO",                                # fixo
    "Área Pública": "Área Pública",                            # coluna DF
    "Motivação": "Motivação",                                  # coluna DF
    "Intenção": "Intenção",                                    # coluna DF
    "Altura (m)": "H",                                         # coluna DF
    "Diâmetro da copa(m)": "Copa",                             # coluna DF
    "DAP 1(cm)": "DAP 1",                                      # coluna DF
    "DAP 2(cm)": "DAP 2",                                      # coluna DF
    "DAP 3(cm)": "DAP 3",                                      # coluna DF
    "DAP 4(cm)": "DAP 4",                                      # coluna DF
    "DAP 5(cm)": "DAP 5",                                      # coluna DF
}

# Campo no site (chave de MAPEAMENTO_PREENCHIMENTO) -> id do elemento/parâmetro no formulário
CAMPO_SITE_PARA_ID_FORM = {
    "Nº no Projeto": "numero_especie_projeto",
    "Nome Popular": "nome_popular",
    "Nome Científico": "nome_cientifico",
    "Observação": "observacao",
    "Estado de Conservação": "estado_conservacao",
    "Local do Espécime": "local_especime",
    "Políticas Municipais": "fcb",
    "Notabilidade": "notabilidade",
    "Utilidade Pública": "utilidade_publica",
    "Área Pública": "area_publica",
    "Motivação": "motivacao",
    "Intenção": "intencao",
    "Altura (m)": "altura_arvore",
    "Diâmetro da copa(m)": "diametro_copa",
    "DAP 1(cm)": "dap1",
    "DAP 2(cm)": "dap2",
    "DAP 3(cm)": "dap3",
    "DAP 4(cm)": "dap4",
    "DAP 5(cm)": "dap5",
}


def obter_valores_mapeamento(row, colunas_df):
    """
    Retorna um dicionário id_form -> valor a enviar, usando MAPEAMENTO_PREENCHIMENTO.
    Se a origem for coluna do DF, usa row[col]; senão usa o valor fixo.
    Nome Popular e Nome Científico ficam como texto (serão convertidos para id depois).
    """
    valores = {}
    for campo_site, origem in MAPEAMENTO_PREENCHIMENTO.items():
        id_form = CAMPO_SITE_PARA_ID_FORM.get(campo_site)
        if not id_form:
            continue
        if origem in colunas_df:
            v = row.get(origem)
            if pd.isna(v):
                v = ""
            else:
                v = str(v).strip()
        else:
            v = "" if origem is None else str(origem).strip()
            
        # --- REGRAS ESPECÍFICAS DE PREENCHIMENTO ---
        if campo_site == "Estado de Conservação":
            val_upper = v.upper()
            if "NÃO ENQUADRADAS" in val_upper:
                v = "Espécime não enquadrada nos casos acima"
            elif "EXÓTICA OU NATIVA, NÃO MA, >=80CM" in val_upper:
                v = "Especies de origem exótica ou nativa não pertencente ao Bioma Mata Atlântica, com DAP >= 80cm"
            elif "NATIVAS MA >= 70CM" in val_upper:
                v = "Espécimes nativas do bioma Mata Atlântica com DAP >= 70cm"
                
        elif campo_site == "Motivação":
            val_upper = v.upper()
            if "SEM MOTIVO" in val_upper:
                v = "SEM MOTIVO"
            elif any(x in val_upper for x in ["MORTA", "QUEBRADA", "CUPIM", "TOMBADA", "PODRE"]):
                v = "MORTE"
            else:
                v = "PROJETO"  # Restante fica como PROJETO
                
        elif campo_site == "Intenção":
            val_upper = v.upper()
            if "PRESERVAR" in val_upper:
                v = "PRESERVAÇÃO"
            elif "REMOVER" in val_upper:
                v = "CORTE"
        # -------------------------------------------

        valores[id_form] = v
    return valores


# Mapeamento texto (planilha) -> value (id) para selects que o servidor só aceita por id
MAPEAMENTO_ESTADO_CONSERVACAO_TEXTO_PARA_VALUE = {
    "8": "8",
    "NÃO ENQUADRADAS": "8",
    "Espécies não enquadradas nos casos acima": "8",
    "ESPÉCIES NÃO ENQUADRADAS": "8",
    "ESPÉCIME NÃO ENQUADRADA NOS CASOS ACIMA": "8",
    "ESPECIES DE ORIGEM EXÓTICA OU NATIVA NÃO PERTENCENTE AO BIOMA MATA ATLÂNTICA, COM DAP >= 80CM": "7",
    "EXÓTICA OU NATIVA, NÃO MA, >=80CM": "7",
    "ESPÉCIMES NATIVAS DO BIOMA MATA ATLÂNTICA COM DAP >= 70CM": "6",
    "NATIVAS MA >= 70CM": "6",
}
MAPEAMENTO_FCB_TEXTO_PARA_VALUE = {
    "3": "3",
    "Espécime não enquadrada nos casos acima": "3",
}
MAPEAMENTO_MOTIVACAO_TEXTO_PARA_VALUE = {
    "1": "1", "PROJETO": "1", "2": "2", "MORTE": "2", "MORTA": "2", "3": "3", "SEM MOTIVO": "3",
    "TERRAPLENAGEM": "1", "REMOVER": "1",
    "QUEBRADA": "2", "CUPIM": "2", "TOMBADA": "2", "PODRE": "2",
}
MAPEAMENTO_INTENCAO_TEXTO_PARA_VALUE = {
    "1": "1", "CORTE": "1", "REMOVER": "1", "2": "2", "PRESERVAÇÃO": "2", "PRESERVAR": "2",
    "3": "3", "TRANSPLANTIO": "3", "4": "4", "AUTORIZAÇÃO ANTERIOR": "4",
}

# Mapeamento planilha → texto exato do select no site (quando difere por grafia/acento/hífen)
# Sibipiruna: igual no site (normalização resolve). Cenostigma sp / samanea sp: ponto após "sp" no site é tratado pela normalização.
NOME_POPULAR_PLANILHA_PARA_SITE = {
    "Figueira Branca": "Figueira-Branca",
    "Aroeirinha": "Aroerinha",
    "Ipê Roxo": "Ipê-rOXO",
    "Árvore samambaia": "árvore-samambaia",
    "Ficus italiano": "ficus-italiano",
    "ficus italiano": "ficus-italiano",
    "Ficus lyrata": "ficus-lira",
    "ficus lyrata": "ficus-lira",
    # Mapeamentos novos solicitados
    "Abacateiro": "Abacate",
    "Areca-bambu": "Areca",
    "Aroeira-vermelha": "Aroerinha",
    "aroeira-vermelha": "aroerinha", 
    "Arvore da chuva": "Samanea",
    "Cassia-rosa": "Cassia",
    "Clusia": "Abaneiro",
    "Eucalipto Citriodora": "Eucalipto",
    "Felicio": "Arvore Samambaia",
    "Ficus Benjamina": "Ficus-Bejamina",
    "Ficus Lirata": "Ficus-Lira",
    "Figueira brava": "Figueira Branca",
    "Figueira Religiosa": "Árvore-do-buda",
    "Figueira-elastica": "Ficus Italiano",
    "Goiabeira": "Goiaba",
    "Ingá do brejo": "Ingá-Banana",
    "Ipê Amarelo": "Ipê Tabaco",
    "Ipê-rOXO": "IPÊ-ROXO",
    "Jerivá": "Baba-de-boi",
    "Morta": "não-identificada",
    "Palmeira Fênix": "Tâmara-mirim",
    "Tapiá": "Tapiá-de-bola",
    "Toco": "não-identificada",
    "toco": "não-identificada",
}
NOME_CIENTIFICO_PLANILHA_PARA_SITE = {
    "Cratateva tapia": "Crataeva tapia",
    "Crataeva tapia": "Crataeva tapia",
    # Mapeamentos novos solicitados
    "Handroanthus avellanedae": "Handroanthus Heptaphyllus",
    "Mimosa caesalpiniifolia": "Mimosa Caesalpiniaefolia",
    "Cenostigma pluviosum": "Cenostigma sp.",
    "Crateva tapia": "Crataeva tapia",
    "Schinus terebinthifolia": "Schinus Terebinthifolius",
    "Corymbia citriodora": "Eucalyptus sp.", 
    "corymbia citriodora": "eucalyptus sp.", 
    "Samanea saman": "Samanea sp.",          
    "samanea saman": "samanea sp.",          
}


def normalizar_payload_requests(payload):
    """
    Ajusta o payload para o formato que o servidor SisArv aceita:
    - numero_especie_projeto: inteiro (64 não 64.0)
    - estado_conservacao, fcb, motivacao, intencao: value (id) do select, não texto
    - altura_arvore, diametro_copa: formato "X,XX" (vírgula)
    - dap1..dap5: inteiro como string
    """
    p = dict(payload)
    if "numero_especie_projeto" in p and p["numero_especie_projeto"]:
        try:
            p["numero_especie_projeto"] = str(int(float(str(p["numero_especie_projeto"]).strip().replace(",", "."))))
        except (ValueError, TypeError):
            pass
    if "estado_conservacao" in p and p["estado_conservacao"] and not str(p["estado_conservacao"]).strip().isdigit():
        v = str(p["estado_conservacao"]).strip().upper()
        for k, id_val in MAPEAMENTO_ESTADO_CONSERVACAO_TEXTO_PARA_VALUE.items():
            if k.upper() in v or v in k.upper():
                p["estado_conservacao"] = id_val
                break
        else:
            p["estado_conservacao"] = "8"
    if "fcb" in p and p["fcb"] and not str(p["fcb"]).strip().isdigit():
        v = str(p["fcb"]).strip()
        p["fcb"] = MAPEAMENTO_FCB_TEXTO_PARA_VALUE.get(v, "3")
    if "motivacao" in p and p["motivacao"] and not str(p["motivacao"]).strip().isdigit():
        v = str(p["motivacao"]).strip().upper()
        p["motivacao"] = MAPEAMENTO_MOTIVACAO_TEXTO_PARA_VALUE.get(v) or "1"
    if "intencao" in p and p["intencao"] and not str(p["intencao"]).strip().isdigit():
        v = str(p["intencao"]).strip().upper()
        p["intencao"] = MAPEAMENTO_INTENCAO_TEXTO_PARA_VALUE.get(v) or "1"
    for campo in ("altura_arvore", "diametro_copa"):
        if campo in p and p[campo] is not None and str(p[campo]).strip():
            try:
                num = float(str(p[campo]).strip().replace(",", "."))
                p[campo] = f"{num:.2f}".replace(".", ",")
            except (ValueError, TypeError):
                pass
    for campo in ("dap1", "dap2", "dap3", "dap4", "dap5"):
        if campo in p and p[campo] is not None and str(p[campo]).strip() != "":
            try:
                p[campo] = str(int(float(str(p[campo]).strip().replace(",", "."))))
            except (ValueError, TypeError):
                p[campo] = "0"
    return p


def normalizar_nome(s):
    """Normaliza nome para comparação: minúsculas, sem acentos, hífens, espaços nem pontos."""
    if not s or not str(s).strip():
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("-", "").replace("–", "").replace("—", "")
    s = s.replace(" ", "")
    s = s.replace(".", "")
    return s


def pausa(min_s=0.5, max_s=1.2):
    """Pausa aleatória para simular comportamento humano."""
    time.sleep(random.uniform(min_s, max_s))


def extrair_numeros_ja_preenchidos(html):
    """Extrai da tabela de árvores do inventário os 'Nº no Projeto' já preenchidos."""
    if not html:
        return set()
    # Tabela no painel de árvores: tbody com <tr> onde a primeira coluna costuma ser o Nº
    m = re.search(
        r'id=["\']?panelArvores["\']?[^>]*>.*?<table[^>]*>.*?<tbody>(.*?)</tbody>',
        html, re.DOTALL | re.IGNORECASE
    )
    if not m:
        return set()
    tbody = m.group(1)
    nums = re.findall(r'<tr[^>]*>\s*<td[^>]*>\s*(\d+)\s*</td>', tbody)
    return {int(x) for x in nums}


def extrair_ids_arvores(html):
    """Extrai os id_inventario_botanico_especie da página (parâmetro de excluiArvore)."""
    if not html:
        return []
    ids = re.findall(r"excluiArvore\s*\(\s*['\"](\d+)['\"]", html)
    return list(dict.fromkeys(ids))  # ordem preservada, sem duplicatas


def preprocessar_df(df):
    """Normaliza o DataFrame para o formato esperado (mesmo layout do Excel do inventário)."""
    df = df.copy()
    if len(df) > 0 and df.index[0] == 0:
        # Remove a linha 2 do cabeçalho mesclado se for a segunda linha de cabeçalho
        df = df.drop(index=0).reset_index(drop=True)
    df = df.rename(columns={"Nome": "Nome Vulgar", "Unnamed: 2": "Nome Científico"})
    df = df.rename(columns={
        "DAP": "DAP 1",
        "Unnamed: 6": "DAP 2",
        "Unnamed: 7": "DAP 3",
        "Unnamed: 8": "DAP 4",
        "Unnamed: 9": "DAP 5",
    })
    for col in ["DAP 1", "DAP 2", "DAP 3", "DAP 4", "DAP 5"]:
        if col in df.columns:
            df[col] = df[col].fillna(0)
    return df


# O servidor pode responder com uma página que redireciona via POST (JavaScript).
def seguir_redirect_post(html, session, max_vezes=5):
    for _ in range(max_vezes):
        if "document.redir.submit()" not in html and len(html) > 500:
            return html
        resp = session.post(f"{base_url}/index.php", data={})
        resp.raise_for_status()
        html = resp.text
    return html


def run_sisarv(formusuario, formsenha, df, progress_callback=None, should_stop=None, progress_range_callback=None):
    """
    Executa o fluxo completo: login no SisArv, exclusão das árvores existentes, inclusão das linhas do df.
    progress_callback(msg) é chamado opcionalmente para atualizar interface (ex.: Streamlit).
    progress_range_callback(atual, total) opcional: chamado a cada árvore (ex.: para barra de progresso).
    should_stop() opcional: se retornar True, interrompe e retorna (False, [], "Interrompido pelo usuário.").
    Retorna: (sucesso: bool, arvores_nao_encontradas: list, mensagem_erro: str|None)
    """
    def stopped():
        return should_stop is not None and should_stop()

    def log(msg):
        if progress_callback:
            progress_callback(msg)
        else:
            print(msg)

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
    }

    session = requests.Session()
    session.headers.update(headers)

    session.get(f"{base_url}/")
    resp_login_page = session.post(
        f"{base_url}/index.php",
        data={"action": "AbreTelaLogin"},
    )
    resp_login_page.raise_for_status()
    html_login = resp_login_page.text

    match_csrf = re.search(r'name="csrf_key"[^>]*value="([^"]+)"', html_login)
    csrf_key = match_csrf.group(1) if match_csrf else ""

    response = session.post(
        f"{base_url}/index.php",
        data={
            "action": "AutenticaUsuario",
            "csrf_key": csrf_key,
            "formusuario": formusuario,
            "formsenha": formsenha,
        },
    )
    response.raise_for_status()
    html = response.text

    html = seguir_redirect_post(html, session)

    response = session.post(
        f"{base_url}/index.php",
        data={"action": "AbreTelaConsultaInventarioBotanico"},
    )
    response.raise_for_status()
    html = response.text
    html = seguir_redirect_post(html, session)

    match_editar = re.search(
        r"abreTelaCadastroInventarioBotanico\s*\(\s*['\"](\d+)['\"]\s*,\s*['\"]consulta['\"]\s*\)",
        html,
    )
    id_inventario = match_editar.group(1) if match_editar else None
    if not id_inventario:
        return (False, [], "Nenhum inventário encontrado na lista para editar.")

    response = session.post(
        f"{base_url}/index.php",
        data={
            "action": "AbreTelaCadastroInventarioBotanico",
            "id_inventario_botanico": id_inventario,
            "origem": "consulta",
        },
    )
    response.raise_for_status()
    html_edicao = response.text
    html_edicao = seguir_redirect_post(html_edicao, session)

    if NAO_PREENCHER:
        if gerar_arquivo_sem_correspondencia:
            gerar_arquivo_sem_correspondencia(df, html_edicao)
        return (True, [], None)

    ids_arvores = extrair_ids_arvores(html_edicao)
    if ids_arvores:
        if stopped():
            return (False, [], "Interrompido pelo usuário.")
        log(f"Excluindo {len(ids_arvores)} árvore(s) do inventário antes de incluir...")
        num_workers = min(4, len(ids_arvores))

        def _excluir_uma(id_esp):
            try:
                resp = session.post(
                    f"{base_url}/index.php",
                    data={
                        "action": "ExcluiArvoreInventarioBotanico",
                        "id_inventario_botanico_especie": id_esp,
                        "origem": "consulta",
                        "id_inventario_botanico": id_inventario,
                    },
                )
                resp.raise_for_status()
                return (id_esp, None)
            except Exception as e:
                return (id_esp, e)

        with ThreadPoolExecutor(max_workers=num_workers) as executor:
            resultados = list(executor.map(_excluir_uma, ids_arvores))

        erros = [(id_esp, err) for id_esp, err in resultados if err is not None]
        if erros:
            for id_esp, err in erros:
                log(f"Erro ao excluir id_inventario_botanico_especie={id_esp}: {err}")
        response = session.post(
            f"{base_url}/index.php",
            data={
                "action": "AbreTelaCadastroInventarioBotanico",
                "id_inventario_botanico": id_inventario,
                "origem": "consulta",
            },
        )
        response.raise_for_status()
        html_edicao = seguir_redirect_post(response.text, session)
        log("Árvores excluídas.")
        if stopped():
            return (False, [], "Interrompido pelo usuário.")

    df_linhas = df.iloc[0:].copy() if len(df) > 0 else pd.DataFrame()
    if df_linhas.empty:
        log("Nenhuma linha no dataframe.")
        return (True, [], None)

    driver = None
    if not USAR_APENAS_REQUESTS and USAR_WEBDRIVER_MANAGER:
        chrome_options = ChromeOptions()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        for tentativa in range(2):
            try:
                service = ChromeService(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)
                break
            except Exception as e:
                if tentativa == 0:
                    log(f"Selenium falhou: {e}. Tentando de novo em 2s...")
                    time.sleep(2)
                else:
                    log("Selenium indisponível. Preenchimento será feito via requests.")
                    break
    elif not USAR_APENAS_REQUESTS:
        log("webdriver-manager não instalado. Preenchimento será feito via requests.")

    if driver is not None:
        driver.set_page_load_timeout(60)
        wait = WebDriverWait(driver, 20)
        url_site = "https://sisarv.rio.gov.br/"
        url_quente = "https://www.google.com"
        try:
            log("Abrindo navegador e carregando página inicial...")
            driver.get(url_quente)
            time.sleep(2.0)
            log("Navegando para o SisArv...")
            driver.get(url_site)
            time.sleep(2.0)
            if driver.current_url in ("data:", "data:,") or "sisarv" not in driver.current_url.lower():
                driver.get(url_site)
                time.sleep(2.0)
            pausa(1.0, 2.0)
            driver.execute_script("document.forms['redir'].submit();")
            pausa(2.0, 3.0)
            log("Fazendo login...")
            wait.until(EC.presence_of_element_located((By.NAME, "formusuario")))
            pausa(0.6, 1.2)
            driver.find_element(By.NAME, "formusuario").clear()
            pausa(0.2, 0.5)
            driver.find_element(By.NAME, "formusuario").send_keys(formusuario)
            pausa(0.4, 0.9)
            driver.find_element(By.NAME, "formsenha").clear()
            pausa(0.2, 0.5)
            driver.find_element(By.NAME, "formsenha").send_keys(formsenha)
            pausa(0.5, 1.0)
            driver.find_element(By.ID, "logForm").submit()
            pausa(3.0, 5.0)
            if "document.redir.submit()" in driver.page_source or len(driver.page_source) < 1000:
                driver.execute_script("document.forms['redir'].submit();")
                pausa(2.0, 3.0)
            log("Indo para Consultar Inventário Botânico...")
            menu_inv = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(.,'Inventário Botânico') and contains(@class,'dropdown-toggle')]"))
            )
            pausa(0.5, 1.0)
            menu_inv.click()
            pausa(0.6, 1.2)
            driver.find_element(By.ID, "opcaoMenu-ConsultarInventarioBotanico").click()
            pausa(2.5, 4.0)
            log("Abrindo tela de Edição do inventário...")
            btn_editar = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@onclick,\"abreTelaCadastroInventarioBotanico\") and contains(@onclick,\"consulta\")]"))
            )
            pausa(0.5, 1.0)
            btn_editar.click()
            pausa(2.5, 4.0)
            log("Preenchendo campo a campo e clicando em Incluir Árvore na Lista...")
            panel_arvores = driver.find_element(By.ID, "panelArvores")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", panel_arvores)
            pausa(1.0, 1.8)
            numeros_ja = extrair_numeros_ja_preenchidos(driver.page_source)
            total_arvores = len(df_linhas)
            pbar = tqdm(df_linhas.iterrows(), total=total_arvores, desc="Unidades", unit="un")
            for idx, (_, row) in enumerate(pbar, start=1):
                if progress_range_callback:
                    progress_range_callback(idx, total_arvores)
                if stopped():
                    log("Interrompido pelo usuário.")
                    return (False, [], "Interrompido pelo usuário.")
                n = row["Nº"]
                if pd.isna(n):
                    continue
                n = int(n)
                pbar.set_postfix(unidade=n)
                if n in numeros_ja:
                    pbar.write(f"Nº {n} já preenchido na lista, pulando.")
                    continue
                pbar.write(f"Preenchendo árvore Nº {n} (campo a campo)...")
                nome_vulgar = str(row["Nome Vulgar"]).strip() if pd.notna(row.get("Nome Vulgar")) else ""
                nome_cientifico = str(row["Nome Científico"]).strip() if pd.notna(row.get("Nome Científico")) else ""
                if not nome_vulgar:
                    nome_vulgar = "não-identificada"
                if not nome_cientifico:
                    nome_cientifico = "ni"
                texto_popular = CORRESPONDENCIAS_NOME_POPULAR.get(nome_vulgar) or nome_vulgar
                texto_cientifico = CORRESPONDENCIAS_NOME_CIENTIFICO.get(nome_cientifico) or nome_cientifico
                texto_popular = NOME_POPULAR_PLANILHA_PARA_SITE.get(texto_popular.strip()) or NOME_POPULAR_PLANILHA_PARA_SITE.get(texto_popular) or texto_popular
                texto_cientifico = NOME_CIENTIFICO_PLANILHA_PARA_SITE.get(texto_cientifico.strip()) or NOME_CIENTIFICO_PLANILHA_PARA_SITE.get(texto_cientifico) or texto_cientifico
                valores = obter_valores_mapeamento(row, df_linhas.columns)
                pausa(0.8, 1.5)
                # Preencher campo a campo (MAPEAMENTO_PREENCHIMENTO): Nº, Nome Popular, Nome Científico, depois demais campos
                # Selects que usam texto visível (nome popular/científico)
                try:
                    Select(driver.find_element(By.ID, "nome_popular")).select_by_visible_text(texto_popular)
                except Exception:
                    try:
                        Select(driver.find_element(By.ID, "nome_popular")).select_by_visible_text(texto_popular.upper())
                    except Exception:
                        pass
                pausa(0.4, 0.9)
                try:
                    Select(driver.find_element(By.ID, "nome_cientifico")).select_by_visible_text(texto_cientifico)
                except Exception:
                    try:
                        Select(driver.find_element(By.ID, "nome_cientifico")).select_by_visible_text(texto_cientifico.upper())
                    except Exception:
                        pass
                pausa(0.3, 0.6)
                # Demais campos a partir do MAPEAMENTO_PREENCHIMENTO
                ids_select = (
                    "estado_conservacao", "local_especime", "fcb",
                    "notabilidade", "utilidade_publica", "area_publica",
                    "motivacao", "intencao",
                )
                for id_form, valor in valores.items():
                    if id_form in ("nome_popular", "nome_cientifico"):
                        continue
                    try:
                        elem = driver.find_element(By.ID, id_form)
                        valor_str = str(valor) if valor else ""
                        if id_form in ids_select:
                            try:
                                Select(elem).select_by_value(valor_str)
                            except Exception:
                                try:
                                    Select(elem).select_by_visible_text(valor_str)
                                except Exception:
                                    pass
                        else:
                            elem.clear()
                            pausa(0.08, 0.2)
                            elem.send_keys(valor_str)
                        pausa(0.1, 0.3)
                    except Exception:
                        pass
                pausa(0.5, 1.0)
                driver.find_element(By.ID, "botao-IncluirArvoreLista").click()
                numeros_ja.add(n)
                pausa(1.5, 2.5)
                pbar.write(f"Nº {n} ({nome_vulgar} / {nome_cientifico}) incluída (navegador).")
            log("Preenchimento da linha 1 ao final concluído (navegador).")
            pausa(2.0, 3.0)
        finally:
            driver.quit()
        return (True, [], None)

    if driver is None:
        # Preenchimento via requests (quando Selenium não está disponível ou falhou)
        def extrair_opcoes_select(html_page, select_id):
            bloco = re.search(
                rf'<select[^>]+id=["\']?{re.escape(select_id)}["\']?[^>]*>(.*?)</select>',
                html_page,
                re.DOTALL | re.IGNORECASE,
            )
            if not bloco:
                return {}
            opts = re.findall(r'<option\s+value="(\d+)"[^>]*>\s*([^<]+?)\s*</option>', bloco.group(1))
            return {texto.strip(): val for val, texto in opts if texto.strip()}

        map_popular = extrair_opcoes_select(html_edicao, "nome_popular")
        map_cientifico = extrair_opcoes_select(html_edicao, "nome_cientifico")
        map_popular_norm = {normalizar_nome(t): val for t, val in map_popular.items()}
        map_cientifico_norm = {normalizar_nome(t): val for t, val in map_cientifico.items()}
        numeros_ja = extrair_numeros_ja_preenchidos(html_edicao)
        arvores_nao_encontradas = []
        total_arvores = len(df_linhas)
        pbar = tqdm(df_linhas.iterrows(), total=total_arvores, desc="Unidades", unit="un")
        for idx, (_, row) in enumerate(pbar, start=1):
            if progress_range_callback:
                progress_range_callback(idx, total_arvores)
            if stopped():
                log("Interrompido pelo usuário.")
                return (False, [], "Interrompido pelo usuário.")
            n = row["Nº"]
            if pd.isna(n):
                continue
            n = int(n)
            pbar.set_postfix(unidade=n)
            if n in numeros_ja:
                pbar.write(f"Nº {n} já preenchido na lista, pulando.")
                continue
            nome_vulgar = str(row["Nome Vulgar"]).strip() if pd.notna(row.get("Nome Vulgar")) else ""
            nome_cientifico = str(row["Nome Científico"]).strip() if pd.notna(row.get("Nome Científico")) else ""
            if not nome_vulgar:
                nome_vulgar = "não-identificada"
            if not nome_cientifico:
                nome_cientifico = "ni"
            texto_popular = CORRESPONDENCIAS_NOME_POPULAR.get(nome_vulgar) or nome_vulgar
            texto_cientifico = CORRESPONDENCIAS_NOME_CIENTIFICO.get(nome_cientifico) or nome_cientifico
            texto_popular = NOME_POPULAR_PLANILHA_PARA_SITE.get(texto_popular.strip()) or NOME_POPULAR_PLANILHA_PARA_SITE.get(texto_popular) or texto_popular
            texto_cientifico = NOME_CIENTIFICO_PLANILHA_PARA_SITE.get(texto_cientifico.strip()) or NOME_CIENTIFICO_PLANILHA_PARA_SITE.get(texto_cientifico) or texto_cientifico
            n_pop = normalizar_nome(texto_popular)
            n_cien = normalizar_nome(texto_cientifico)
            id_popular = (
                map_popular_norm.get(n_pop)
                or map_popular.get(texto_popular)
                or map_popular.get(texto_popular.upper())
            )
            id_cientifico = (
                map_cientifico_norm.get(n_cien)
                or map_cientifico.get(texto_cientifico)
                or map_cientifico.get(texto_cientifico.upper())
            )
            if not id_popular or not id_cientifico:
                arvores_nao_encontradas.append((n, texto_popular, texto_cientifico))
                pbar.write(f"Nº {n}: nome não encontrado nos selects (vulgar={texto_popular!r}, científico={texto_cientifico!r}). Pulando.")
                continue
            valores = obter_valores_mapeamento(row, df_linhas.columns)
            valores["nome_popular"] = id_popular
            valores["nome_cientifico"] = id_cientifico
            payload = {
                "action": "IncluiArvoreInventarioBotanico",
                "id_inventario_botanico": id_inventario,
                "origem": "consulta",
                "id_em_edicao": "",
                "area_interesse_social": "SIM",
                **valores,
            }
            payload = normalizar_payload_requests(payload)
            resp = session.post(f"{base_url}/index.php", data=payload)
            try:
                resp.raise_for_status()
            except requests.exceptions.HTTPError as e:
                pbar.write(f"Nº {n}: servidor retornou {resp.status_code} - {e}")
                pbar.write(f"Resposta: len={len(resp.text)} chars; primeiros 800: {repr(resp.text[:800])}")
                pbar.write("Payload (valores enviados):")
                for k, v in payload.items():
                    if k == "action":
                        continue
                    pbar.write(f"  {k}={repr(v)}")
                pbar.write("Pulando para a próxima árvore.")
                continue
            html_edicao = seguir_redirect_post(resp.text, session)
            numeros_ja = extrair_numeros_ja_preenchidos(html_edicao)
            pbar.write(f"Nº {n} ({nome_vulgar} / {nome_cientifico}) incluída via requests.")
        log("Preenchimento da linha 1 ao final concluído (via requests).")
        if arvores_nao_encontradas:
            log("--- Árvores não encontradas nos selects ---")
            for n, vulg, cien in arvores_nao_encontradas:
                log(f"  Nº {n}: {vulg!r} / {cien!r}")
            log(f"Total: {len(arvores_nao_encontradas)} árvore(s) não encontrada(s).")
        return (True, arvores_nao_encontradas, None)


if __name__ == "__main__":
    caminho_excel = r"C:\Users\DE0189769\OneDrive - Direcional Engenharia S A\Documentos Macedo One Drive\Automações - Lucas\ws.py"
    df = pd.read_excel(caminho_excel)
    df = preprocessar_df(df)
    print(df)
    run_sisarv("evelin.caboclo@gmail.com", "67750131", df)
