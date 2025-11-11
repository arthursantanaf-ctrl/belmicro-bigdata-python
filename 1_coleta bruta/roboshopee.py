import pandas as pd
import time
import undetected_chromedriver as uc
import random
import re # Importamos a biblioteca de expressões regulares para limpeza de texto
from urllib.parse import quote
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# --- CONFIGURAÇÕES GERAIS ---
ARQUIVO_ENTRADA = r'C:\Users\asf\Documents\resultado final shopee\coleta bruta\lista produtos.xlsx' 
ARQUIVO_SAIDA = r'C:\Users\asf\Documents\resultado final shopee\coleta bruta\resultados_shopee_finalissimo.xlsx'
NOME_COLUNA_PESQUISA = 'Descricao'
MAX_PRODUTOS_POR_PESQUISA = 45

def configurar_driver():
    """Configura o Chrome usando o undetected-chromedriver com versão especificada."""
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    
    caminho_perfil_dedicado = r'C:\meu-perfil-selenium' 
    options.add_argument(f'--user-data-dir={caminho_perfil_dedicado}')
    
    print("Iniciando driver com undetected-chromedriver...")
    
    # Verifique sua versão em Ajuda > Sobre o Google Chrome
    versao_do_chrome = 142
    driver = uc.Chrome(options=options, use_subprocess=True, version_main=versao_do_chrome) 
    return driver

def limpar_numero(texto):
    """Função para limpar e converter texto em número (int ou float)."""
    if not isinstance(texto, str):
        return texto
    
    texto_limpo = texto.lower()
    # Converte "mil" em "000" e remove "k" (assumindo que "k" também significa mil)
    if 'mil' in texto_limpo:
        texto_limpo = texto_limpo.replace('mil', '000')
    if 'k' in texto_limpo:
        texto_limpo = texto_limpo.replace('k', '000')

    # Remove todos os caracteres não numéricos, exceto a vírgula
    numeros = re.sub(r'[^\d,]', '', texto_limpo)
    
    # Se houver vírgula, substitui por ponto para converter para float
    if ',' in numeros:
        numeros = numeros.replace(',', '.')
        try:
            return float(numeros)
        except ValueError:
            return texto # Retorna o texto original se a conversão falhar
    else:
        try:
            return int(numeros)
        except ValueError:
            return texto # Retorna o texto original se a conversão falhar

def extrair_dados_produto(driver, url_produto):
    driver.get(url_produto)
    dados_produto = {
        'Nome': 'Não encontrado',
        'Preço (R$)': 'Não encontrado',
        'Avaliação Média': 'Não encontrado',
        'Total de Avaliações': 'Não encontrado',
        'Vendidos': 'Não encontrado',
        'Vendedor': 'Não encontrado',
        'Link Loja': 'Não encontrado',
        'URL': url_produto
    }

    wait = WebDriverWait(driver, 15)

    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.page-product')))
        time.sleep(random.uniform(1.5, 2.5)) # Pausa para JS carregar

        # ------------------------
        # Nome
        # ------------------------
        try:
            dados_produto['Nome'] = driver.find_element(By.XPATH, '//h1').text.strip()
        except:
            pass

        # ------------------------
        # Preço
        # ------------------------
        preco = None
        for xpath in [
            '//div[contains(@class,"IZPeQz")]',
            '//div[contains(@class,"pqTWkA")]',
            '//span[contains(text(),"R$")]'
        ]:
            try:
                preco = driver.find_element(By.XPATH, xpath).text
                break
            except:
                continue
        if preco:
            preco = preco.replace("R$", "").strip()
            dados_produto['Preço (R$)'] = limpar_numero(preco)

        # ------------------------
        # Avaliação média
        # ------------------------
        avaliacao = None
        for xpath in [
            '(//button[contains(@class,"e2p50f")]/div)[1]',
            '//div[contains(@class,"product-rating-overview__rating-score")]'
        ]:
            try:
                avaliacao = driver.find_element(By.XPATH, xpath).text
                break
            except:
                continue
        if avaliacao:
            dados_produto['Avaliação Média'] = float(avaliacao.replace(",", ".").strip())

        # ------------------------
        # Total de avaliações
        # ------------------------
        total_av = None
        for xpath in [
            '//button[contains(@class,"e2p50f")]/div[@class="F9RHbS"]',
            '//div[contains(text(),"avaliações")]'
        ]:
            try:
                total_av = driver.find_element(By.XPATH, xpath).text
                break
            except:
                continue
        if total_av:
            dados_produto['Total de Avaliações'] = limpar_numero(total_av)

        # ------------------------
        # Vendidos
        # ------------------------
        try:
            vendidos_elem = driver.find_element(By.CSS_SELECTOR, "div.aleSBU")
            vendidos_texto = vendidos_elem.text.strip()
            dados_produto['Vendidos'] = limpar_numero(vendidos_texto)
        except NoSuchElementException:
            pass # Silencioso, "Não encontrado" é o padrão

        # ------------------------
        # Vendedor (Nome + Link da Loja)
        # ------------------------
        try:
            vendedor_nome = driver.find_element(By.CSS_SELECTOR, "section.page-product__shop div.fV3TIn").text
            dados_produto['Vendedor'] = vendedor_nome.strip()
        except NoSuchElementException:
            pass
        try:
            vendedor_link = driver.find_element(By.CSS_SELECTOR, "section.page-product__shop a.lG5Xxv").get_attribute("href")
            dados_produto['Link Loja'] = vendedor_link
        except NoSuchElementException:
            pass

    except TimeoutException:
        print(f"⏳ Timeout ao carregar: {url_produto}")

    return dados_produto

def main():
    print("Iniciando o processo de scraping da Shopee...")
    driver = configurar_driver()
    wait = WebDriverWait(driver, 15)
    
    driver.get("https://shopee.com.br/")
    print("\n" + "="*80)
    input("### AÇÃO NECESSÁRIA: Se for o primeiro uso, faça o login na Shopee. ###\n### Depois, volte aqui e pressione Enter para iniciar a pesquisa. ###")
    print("="*80 + "\n")

    pausa_inicial = random.uniform(3, 5)
    print(f"Ok, aguardando {pausa_inicial:.1f} segundos antes de começar...")
    time.sleep(pausa_inicial)

    try:
        df_pesquisas = pd.read_excel(ARQUIVO_ENTRADA)
        print(f"Planilha '{ARQUIVO_ENTRADA}' lida com sucesso. {len(df_pesquisas)} itens para pesquisar.")
    except FileNotFoundError:
        print(f"ERRO: O arquivo '{ARQUIVO_ENTRADA}' não foi encontrado. Verifique o caminho no código.")
        driver.quit()
        return
        
    todos_os_dados = []
    
    for index, linha in df_pesquisas.iterrows():
        termo_pesquisa = linha[NOME_COLUNA_PESQUISA]
        if pd.isna(termo_pesquisa): continue

        print(f"\n[{index + 1}/{len(df_pesquisas)}] Pesquisando por: '{termo_pesquisa}'")
        
        # --- LÓGICA DE PAGINAÇÃO v1.5 ---
        
        urls_para_visitar_total = [] # Lista de links para este termo
        numero_pagina = 0 # Começa na página 1 (que tem o índice 0)
        limite_paginas = 5 # Um limite de segurança para não rodar para sempre
        
        while len(urls_para_visitar_total) < MAX_PRODUTOS_POR_PESQUISA and numero_pagina < limite_paginas:
            
            print(f"  -> Acessando Página {numero_pagina + 1}...")
            
            try:
                termo_formatado = quote(termo_pesquisa)
                # Adicionamos o parâmetro &page={numero_pagina}
                url_de_busca = f"https://shopee.com.br/search?keyword={termo_formatado}&page={numero_pagina}"
                driver.get(url_de_busca)

                seletor_produto = 'li.shopee-search-item-result__item'
                wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, seletor_produto)))
                
                # Pausa para garantir que todos os elementos da página carregaram
                time.sleep(random.uniform(2.0, 3.5))

                print(f"  -> Coletando links da Página {numero_pagina + 1}...")
                
                seletor_links = "li.shopee-search-item-result__item a[href]"
                elementos_link = driver.find_elements(By.CSS_SELECTOR, seletor_links) 
                
                links_desta_pagina = []
                for link in elementos_link:
                    try:
                        href = link.get_attribute('href')
                        
                        # --- MUDANÇA v1.6 (NOVO FILTRO) ---
                        # Adicionamos a condição 'find_similar_products' not in href
                        filtro_1 = href and 'shopee.com.br' in href
                        filtro_2 = 'search' not in href
                        filtro_3 = 'find_similar_products' not in href # <-- NOVO FILTRO AQUI
                        filtro_4 = href not in urls_para_visitar_total
                        
                        if filtro_1 and filtro_2 and filtro_3 and filtro_4:
                            links_desta_pagina.append(href)
                            urls_para_visitar_total.append(href)
                    except:
                        continue 

                # Se a página não retornar nenhum link novo, paramos
                if not links_desta_pagina:
                    print(f"  -> Nenhum link novo encontrado na Página {numero_pagina + 1}. Provavelmente chegamos ao fim.")
                    break
                
                print(f"  -> {len(links_desta_pagina)} links novos encontrados.")
                print(f"  -> Total de links acumulados: {len(urls_para_visitar_total)} (Meta: {MAX_PRODUTOS_POR_PESQUISA})")

                numero_pagina += 1 # Prepara para a próxima página

            except (NoSuchElementException, TimeoutException):
                print(f"  -> Nenhum resultado encontrado na Página {numero_pagina + 1}. Parando a busca por este termo.")
                break # Para o loop 'while' e vai para o próximo termo

        # --- FIM DO BLOCO DE PAGINAÇÃO ---

        # Aplicamos o limite MÁXIMO
        urls_para_processar = urls_para_visitar_total[:MAX_PRODUTOS_POR_PESQUISA]

        print(f"\n  -> Busca por '{termo_pesquisa}' concluída.")
        print(f"  -> Total de links válidos encontrados: {len(urls_para_visitar_total)}")
        print(f"  -> Produtos que serão extraídos (limite de {MAX_PRODUTOS_POR_PESQUISA}): {len(urls_para_processar)}")
        
        if not urls_para_processar:
            print("  -> Nenhum link válido encontrado para este termo. Pulando.")
            continue

        for i, url in enumerate(urls_para_processar):
            print(f"    - Extraindo dados [{i+1}/{len(urls_para_processar)}]: {url[:60]}...")
            dados = extrair_dados_produto(driver, url)
            dados['Termo Pesquisado'] = termo_pesquisa
            todos_os_dados.append(dados)
            
            # 💾 Salvamento automático a cada 10 produtos
            if len(todos_os_dados) % 10 == 0:
                df_temp = pd.DataFrame(todos_os_dados)
                df_temp.to_excel(ARQUIVO_SAIDA, index=False)
                print(f"💾 Progresso salvo automaticamente! Total de {len(todos_os_dados)} produtos até agora.")
            
            # Pausa curta entre cada produto
            time.sleep(random.uniform(1.5, 3.0))
            
        # 🕐 Pausa longa a cada X pesquisas
        if (index + 1) % 5 == 0:
            pausa_longa = random.uniform(10, 15)
            print(f"⏸️ Pausa longa de {pausa_longa:.1f}s para evitar CAPTCHA...")
            time.sleep(pausa_longa)

            
    if todos_os_dados:
        df_resultados = pd.DataFrame(todos_os_dados)
        df_resultados.to_excel(ARQUIVO_SAIDA, index=False)
        print(f"\nProcesso finalizado! Os dados foram salvos em '{ARQUIVO_SAIDA}'.")
    else:
        print("\nNenhum dado foi coletado. O arquivo de saída não foi gerado.")
    driver.quit()

if __name__ == "__main__":
    main()