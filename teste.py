import asyncio
import sys
from playwright.async_api import async_playwright
import datetime as dt # Importar datetime
import os

# --- lógica de cálculo de datas ---
date_today = dt.date.today()
tresmesesatras = date_today.fromordinal(date_today.toordinal()-92).strftime('%d/%m/%Y')
seismesesatras = date_today.fromordinal(date_today.toordinal()-184).strftime('%d/%m/%Y')
dozemesesatras = date_today.fromordinal(date_today.toordinal()-365).strftime('%d/%m/%Y')
hoje_formatado = date_today.strftime('%d/%m/%Y')

#criar um dicionário com as datas
datas_formatadas = {
    'hoje': hoje_formatado,
    'doze_meses_atras': dozemesesatras,
    'seis_meses_atras': seismesesatras,
    'tres_meses_atras': tresmesesatras
}
print(f"INFO: Data de hoje: {hoje_formatado}") 
print(f"INFO: Doze meses atrás: {dozemesesatras}")
print(f"INFO: Seis meses atrás: {seismesesatras}")
print(f"INFO: Três meses atrás: {tresmesesatras}")
# --- Fim da lógica de datas ---


# --- Fazendo login ---
async def fazer_login(page, login_url, usuario, senha):
    print(f"INFO: Navegando para URL de login: {login_url}")
    await page.goto(login_url, wait_until="networkidle", timeout=60000)
    print("INFO: Preenchendo formulário de login...")
    await page.fill("#username", usuario)
    await page.fill("#current-password", senha)
    await page.click("#login")
    print("INFO: Login submetido.")
    await page.wait_for_selector("#acao_alterar_local", timeout=30000) # Espera elemento pós-login
    print("INFO: Login confirmado.")

# --- Selecionar filial padrão Filial 1---
async def selecionar_filial_padrao(page): # Nome mais genérico se for sempre a mesma
    print("INFO: Iniciando seleção da filial padrão...")
    await page.click("#acao_alterar_local")
    print("INFO: Clicado em 'Alterar Local'.")
    
    await page.locator("it-filial-dynacombobox").get_by_role("button", name="Mostra lista").click()
    print("INFO: Dropdown de filiais aberto.")
    await asyncio.sleep(1) # Pausa para dropdown carregar (use asyncio.sleep para pausas genéricas)
    
    # Usando o seletor que funcionou para você
    await page.get_by_role("option", name="Nome: TRIMAF - TRIANGULO MATERIAIS GRAFICOS EIRELI Código: 1 CPF/CNPJ: 86.518.").click()
    print("INFO: Filial TRIMAF selecionada no dropdown.")
    await asyncio.sleep(0.5)
    
    await page.get_by_role("button", name="Gravar").click()
    print("INFO: Botão 'Gravar' filial clicado.")
    await page.wait_for_load_state('networkidle', timeout=30000) # Espera ação concluir
    print("INFO: Seleção de filial padrão concluída.")
    return True


# --- Fim da seleção de filial padrão ---

# --- Mudar para filial específica ---
async def mudar_para_filial_especifica(page, nome_exato_da_opcao_filial_no_dropdown):
    print(f"\nINFO: Tentando mudar para a filial: '{nome_exato_da_opcao_filial_no_dropdown}'...")
    try:
        await page.click("#acao_alterar_local")
        print("INFO: Clicado em 'Alterar Local'.")
        
        await page.locator("it-filial-dynacombobox").get_by_role("button", name="Mostra lista").click()
        print("INFO: Dropdown de filiais aberto.")
        await asyncio.sleep(1) # Pausa para dropdown carregar
        
        # Usando o seletor que funcionou para você
        await page.get_by_role("option", name=nome_exato_da_opcao_filial_no_dropdown).click()
        print("INFO: Filial TRIMAF selecionada no dropdown.")
        await asyncio.sleep(0.5)
        
        await page.get_by_role("button", name="Gravar").click()
        print("INFO: Botão 'Gravar' filial clicado.")
        await page.wait_for_load_state('networkidle', timeout=30000) # Espera ação concluir
        print(f"INFO: Seleção da filial '{nome_exato_da_opcao_filial_no_dropdown}' concluída.")
        return True # Indica que a mudança foi bem-sucedida
    except Exception as e:
        print(f"ERRO ao tentar mudar para a filial '{nome_exato_da_opcao_filial_no_dropdown}': {e}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção (Mudança para Filial: {nome_exato_da_opcao_filial_no_dropdown})...")
            await page.pause()
        return False # Indica que a mudança falhou

async def abrir_menu_cadastros(page):
    SELETOR_MENU = "#menucadastro" # Já usamos este, parece correto
    print(f"INFO: Abrindo menu 'Cadastros' (usando seletor: {SELETOR_MENU})...")
    try:
        await page.click(SELETOR_MENU)
        await asyncio.sleep(0.75) # Pequena pausa para o menu/submenu aparecer
        print(f"INFO: Menu 'Cadastros' clicado.")
    except Exception as e:
        print(f"ERRO ao tentar abrir o menu 'Cadastros' ({SELETOR_MENU}): {e}")
        await page.pause() # Pausa para inspeção em caso de erro
        raise # Re-levanta a exceção para parar a lógica dependente

async def abrir_menu_compras(page):
    SELETOR_MENU = "#menucompras" # !! VERIFIQUE ESTE SELETOR COM O INSPECTOR !!
    print(f"INFO: Abrindo menu 'Compras' (usando seletor: {SELETOR_MENU})...")
    try:
        await page.click(SELETOR_MENU)
        await asyncio.sleep(0.75)
        print(f"INFO: Menu 'Compras' clicado.")
    except Exception as e:
        print(f"ERRO ao tentar abrir o menu 'Compras' ({SELETOR_MENU}): {e}")
        await page.pause()
        raise

async def abrir_menu_vendas(page):
    SELETOR_MENU = "#menuvendas" # Já usamos este
    print(f"INFO: Abrindo menu 'Vendas' (usando seletor: {SELETOR_MENU})...")
    try:
        await page.click(SELETOR_MENU)
        await asyncio.sleep(0.75)
        print(f"INFO: Menu 'Vendas' clicado.")
    except Exception as e:
        print(f"ERRO ao tentar abrir o menu 'Vendas' ({SELETOR_MENU}): {e}")
        await page.pause()
        raise

async def abrir_menu_fiscal(page):
    SELETOR_MENU = "#menufiscal" # Já usamos este
    print(f"INFO: Abrindo menu 'Fiscal' (usando seletor: {SELETOR_MENU})...")
    try:
        await page.click(SELETOR_MENU)
        await asyncio.sleep(0.75)
        print(f"INFO: Menu 'Fiscal' clicado.")
    except Exception as e:
        print(f"ERRO ao tentar abrir o menu 'Fiscal' ({SELETOR_MENU}): {e}")
        await page.pause()
        raise

async def abrir_menu_estoque(page):
    SELETOR_MENU = "#menuestoque" # Já usamos este
    print(f"INFO: Abrindo menu 'Estoque' (usando seletor: {SELETOR_MENU})...")
    try:
        await page.click(SELETOR_MENU)
        await asyncio.sleep(0.75)
        print(f"INFO: Menu 'Estoque' clicado.")
    except Exception as e:
        print(f"ERRO ao tentar abrir o menu 'Estoque' ({SELETOR_MENU}): {e}")
        await page.pause()
        raise

async def abrir_menu_financeiro(page):
    SELETOR_MENU = "#menufinanceiro" # !! VERIFIQUE ESTE SELETOR COM O INSPECTOR !!
    print(f"INFO: Abrindo menu 'Financeiro' (usando seletor: {SELETOR_MENU})...")
    try:
        await page.click(SELETOR_MENU)
        await asyncio.sleep(0.75)
        print(f"INFO: Menu 'Financeiro' clicado.")
    except Exception as e:
        print(f"ERRO ao tentar abrir o menu 'Financeiro' ({SELETOR_MENU}): {e}")
        await page.pause()
        raise

async def abrir_menu_painel_contador(page):
    SELETOR_MENU = "#menupainelcontador" # !! VERIFIQUE ESTE SELETOR (ou similar, ex: #menupainelcontabil) !!
    print(f"INFO: Abrindo menu 'Painel Contador' (usando seletor: {SELETOR_MENU})...")
    try:
        await page.click(SELETOR_MENU)
        await asyncio.sleep(0.75)
        print(f"INFO: Menu 'Painel Contador' clicado.")
    except Exception as e:
        print(f"ERRO ao tentar abrir o menu 'Painel Contador' ({SELETOR_MENU}): {e}")
        await page.pause()
        raise

async def abrir_menu_ferramentas(page):
    SELETOR_MENU = "#menuferramentas" # !! VERIFIQUE ESTE SELETOR COM O INSPECTOR !!
    print(f"INFO: Abrindo menu 'Ferramentas' (usando seletor: {SELETOR_MENU})...")
    try:
        await page.click(SELETOR_MENU)
        await asyncio.sleep(0.75)
        print(f"INFO: Menu 'Ferramentas' clicado.")
    except Exception as e:
        print(f"ERRO ao tentar abrir o menu 'Ferramentas' ({SELETOR_MENU}): {e}")
        await page.pause()
        raise






# --- Baixar relatório de Produtos ---

async def baixar_relatorio_produtos(page, datas, id_filial_para_nome_arquivo):
    print("\nINFO: Iniciando processo para baixar relatório de Produtos...") # Adicionado \n para espaçamento
    hoje_formatado_local = datas['hoje']

    try:
        print("INFO: Abrindo menu cadastro.") # Mudado de PASSO para INFO para consistência
        await abrir_menu_cadastros(page)
        print("INFO: Menu cadastro aberto.")
        # Removido wait_for_load_state aqui, pode ser desnecessário se o próximo clique for rápido
        await asyncio.sleep(0.5) # Mantida a pausa original para garantir que o menu esteja pronto
        
        print("INFO: Clicando em Produtos.")
        await page.click("#produtocrudcontroller")
        await page.wait_for_load_state('networkidle', timeout=45000) 
        print("INFO: Página de Produtos aberta.")

        print("PASSO: Clicando em 'Selecionar visão'...")
        await page.click("#edicao_visao_grid")
        print(f"INFO: Clicado em: #edicao_visao_grid")

        print("PASSO: Selecionando a visão desejada no popup...")
        # Supondo que o clique em #edicao_visao_grid abre um popup que contém este botão
        await page.get_by_role("button", name="Mostra lista").click()
        await asyncio.sleep(1)
        await page.get_by_role("option", name="Visão Relatórios").click()
        await asyncio.sleep(0.5)
        print("INFO: Seleção de visão concluída.")

        print("PASSO: Aplicando a visão...")
        await page.click("#btn_aplicar_visao")
        print(f"INFO: Clicado em: #btn_aplicar_visao")
        await page.wait_for_load_state('networkidle', timeout=45000)
        print("INFO: Aplicação de visão concluída.")

        print("PASSO: Clicando em 'Exportar Dados'...")
        await page.click("#exportar_dados")
        print(f"INFO: Clicado em: #exportar_dados")


        print("PASSO: Selecionando formato Excel e XLSX no popup...")
        await asyncio.sleep(1)
        await page.locator("label").filter(has_text="EXCEL").nth(1).click()
        await page.locator("label").filter(has_text="xlsx").nth(1).click()
        await asyncio.sleep(2)

        print("PASSO: Clicando em 'Baixar' e aguardando download...")
        async with page.expect_download(timeout=60000) as download_info:
            await page.get_by_role("button", name=" EXPORTAR").click()
        
        download = await download_info.value
        nome_arquivo_produtos = f"Relatorio_ERP_Produtos_OFICIAL_{id_filial_para_nome_arquivo}.xlsx" 
        caminho_pasta_downloads = datas["hoje"].replace('/', '-')
        if not os.path.exists(caminho_pasta_downloads):
            os.makedirs(caminho_pasta_downloads)
        caminho_para_salvar = f"{caminho_pasta_downloads}/{nome_arquivo_produtos}"
        await download.save_as(caminho_para_salvar)
        print(f"INFO: Relatório de produtos para '{id_filial_para_nome_arquivo}' salvo como: {caminho_para_salvar}")
        print(f"INFO: Processo para baixar relatório de Produtos para '{id_filial_para_nome_arquivo}' concluído.")
    except Exception as e_baixa_produto:
        print(f"ERRO ao baixar relatório de produtos para '{id_filial_para_nome_arquivo}': {e_baixa_produto}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção (Produtos - {id_filial_para_nome_arquivo})...")
            await page.pause()







# --- Baixar relatório de Giro de Estoque ---
async def baixar_relatorio_giro_estoque(page, datas):
        print("INFO: Iniciando processo para baixar relatório de Estoque...")
        hoje_formatado_local = datas["hoje"]
        try:
            # 1. Clicar em menu "estoque"
            print("PASSO: Abrindo menu Estoque.")
            await abrir_menu_estoque(page)

            # 2. Clicar em "giro de estoque"
            print("PASSO: Clicando em Giro de Estoque.")
            SELETOR_GIRO_ESTOQUE_LINK = "#giroestoquereportcontroller" # Exemplo - Verifique
            await page.click(SELETOR_GIRO_ESTOQUE_LINK)
            print(f"INFO: Clicado em: {SELETOR_GIRO_ESTOQUE_LINK}")
            await page.wait_for_load_state('networkidle', timeout=45000) # Esperar página do relatório carregar

            # 3. Preencher o ano
            print("PASSO: Preenchendo o ano (2025).")
            SELETOR_CAMPO_ANO_GIRO = "#Report_ano" # Exemplo - Verifique ID ou use get_by_label
            await page.fill(SELETOR_CAMPO_ANO_GIRO, "2025") 
            print(f"INFO: Ano preenchido em: {SELETOR_CAMPO_ANO_GIRO}")

            print("PASSO: Abrindo dropdown de filiais para Giro de Estoque.")
            await page.locator("wj-multi-select").get_by_role("button", name="Mostra lista").click()
            print(f"INFO: Clicado em: Selecionar filiais para Giro de Estoque.")
            await page.wait_for_timeout(1000) # Esperar dropdown abrir completamente

            print("PASSO: Selecionando Selecionar todas no dropdown de Giro de Estoque.")
            await page.get_by_role("checkbox", name="Selecionar Todos").check()
            print(f"INFO: Filial 1 selecionada (checkbox).")
        
            # Depois de selecionar os checkboxes
            await page.locator("wj-multi-select").get_by_role("button", name="Mostra lista").click()
            print(f"INFO: Fechando dropdown de filiais para Giro de Estoque.")  
            await page.wait_for_timeout(1000)

            # 5. Selecionar formato Excel/XLSX
            print("PASSO: Selecionando formato Excel e XLSX para Giro de Estoque.")  
            await page.locator("label").filter(has_text="EXCEL").first.click()

            await page.locator("label").filter(has_text="xlsx").first.click()
            print("INFO: Formatos de exportação selecionados.")
            await page.wait_for_timeout(1000)

            # 6. Clicar em "Baixar" (ou "Agendar" como no seu XPath)
            print("PASSO: Clicando em 'Baixar/Agendar' para Giro de Estoque e aguardando download...")
        
            timeout_download_giro_ms = 300 * 100000
            async with page.expect_download(timeout=timeout_download_giro_ms) as download_info_giro:
                await page.get_by_role("button", name="b Gerar").click()
        
            download_giro = await download_info_giro.value
            nome_arquivo_giro = f"Relatorio_ERP_RelatÃ³rio - Giro de estoque 2025_OFICIAL.xlsx"
            caminho_pasta_downloads = hoje_formatado_local.replace('/', '-') # Reutilizando a pasta da data

            caminho_para_salvar_giro = f"{caminho_pasta_downloads}/{nome_arquivo_giro}"
            await download_giro.save_as(caminho_para_salvar_giro)
            print(f"INFO: Relatório de Giro de Estoque salvo como: {caminho_para_salvar_giro}")
        
            print("INFO: Processo para baixar relatório de Giro de Estoque concluído.")

        except Exception as e_baixa_giro:
            print(f"ERRO ao baixar relatório de Giro de Estoque: {e_baixa_giro}")
            import traceback
            traceback.print_exc()
            if page and not page.is_closed():
                print("ERRO: Pausando para inspeção (Giro de Estoque)...")
                await page.pause()

        # --- FIM BAIXAR GIRO DE ESTOQUE ---










# --- funçao para baixar relatório de Clientes ---
async def baixar_relatorio_clientes(page, datas):
    print("\nINFO: Iniciando processo para baixar relatório de Clientes...")
    try:
        # 1. Navegação para Clientes
        print("PASSO: Abrindo menu Cadastro.")
        await abrir_menu_cadastros(page) # Usando a função de menu

        print("PASSO: Clicando em Clientes.")
        SELETOR_CLIENTES_LINK = "#clientecrudcontroller" # Verifique se este é o ID correto
        await page.click(SELETOR_CLIENTES_LINK)
        await page.wait_for_load_state('networkidle', timeout=30000)
        print("INFO: Página de Clientes carregada.")

        # 2. Clicar em "Exportar Dados"
        await page.wait_for_load_state('networkidle', timeout=45000)
        print("INFO: Aplicação de visão concluída.")

        print("PASSO: Clicando em 'Exportar Dados'...")
        await asyncio.sleep(2)
        await page.click("#exportar_dados")
        await asyncio.sleep(2)
        print('Clicando novamente em "Exportar Dados" para confirmar...')


        await page.wait_for_load_state('networkidle', timeout=45000)
        # 3. Selecionar formato Excel/XLSX
        
        # >>> NOVA ESPERA ADICIONADA AQUI <<<
        print("INFO: Aguardando popup de exportação de clientes aparecer (pelo título)...")
        # Vamos esperar pelo título do popup que vemos na sua imagem

        print("PASSO: Selecionando formato Excel e XLSX no popup...")
        await page.locator("label").filter(has_text="EXCEL").nth(1).click()
        await page.locator("label").filter(has_text="xlsx").nth(1).click()
        await asyncio.sleep(2)
        
        print("PASSO: Clicando em 'Baixar' e aguardando download...")
        async with page.expect_download(timeout=60000) as download_info:
            await page.get_by_role("button", name=" EXPORTAR").click()
        
        download = await download_info.value
        nome_arquivo_clientes = f"Relatorio_ERP_Clientes_OFICIAL.xlsx" # Você pode adicionar a data aqui se quiser diferenciar
        caminho_pasta_downloads = datas["hoje"].replace('/', '-')
        
        if not os.path.exists(caminho_pasta_downloads):
            os.makedirs(caminho_pasta_downloads)
            
        caminho_para_salvar = f"{caminho_pasta_downloads}/{nome_arquivo_clientes}"
        await download.save_as(caminho_para_salvar)
        print(f"INFO: Relatório de clientes salvo como: {caminho_para_salvar}")
        
        print("INFO: Processo para baixar relatório de Produtos concluído.")


    except Exception as e_clientes:
        print(f"ERRO ao baixar relatório de clientes: {e_clientes}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print("ERRO: Pausando para inspeção (Clientes)...")
            await page.pause()









async def baixar_orcamento_faturamento(page, datas, id_filial_para_nome_arquivo):
    print("\nINFO: Iniciando processo para baixar relatório de Orçamento de Faturamento...")
    hoje_formatado_local = datas["hoje"]
    try:
        # 1. Navegação
        await abrir_menu_vendas(page) # Reutilizando nossa função de menu!
        
        print("PASSO: Clicando em 'Orçamentos de faturamento'.")
        # await page.pause() # Use para encontrar este seletor
        SELETOR_LINK_ORCAMENTO_FATURAMENTO = "#orcamentofaturamentocrudcontroller" # VERIFIQUE este ID!
        await page.click(SELETOR_LINK_ORCAMENTO_FATURAMENTO)
        await page.wait_for_load_state('networkidle', timeout=30000)
        print("INFO: Página de Orçamento de Faturamento carregada.")

        # 2. Aplicar Filtros
        await asyncio.sleep(2)
        print("PASSO: Clicando no botão 'Filtrar' da grid...")
        #await page.pause()
        SELETOR_BTN_FILTRAR_GRID = "#filtrarGrid" # VERIFIQUE este ID!
        await page.click(SELETOR_BTN_FILTRAR_GRID)
        
        # Esperar o popup de filtro aparecer. Use o Inspector para encontrar um bom seletor para o popup.
        # Pode ser um título, um ID do container do popup, etc.
        print("INFO: Aguardando popup de filtro...")

        await page.locator("#FiltroFaturamentoWrapper_status").wait_for(state="visible", timeout=10000)
        print("INFO: Popup de filtro aberto.")
        await asyncio.sleep(1) # Pausa para o popup carregar completamente

        # Aplicar filtros
        print("PASSO: Aplicando filtros...")
        print("PASSO: Escrevendo 'Todos' no combobox do filtro...")
        await page.locator("#FiltroFaturamentoWrapper_status").fill("Todos")
        print('Todos escritos no combobox do filtro.')
        await asyncio.sleep(0.5)
        

        print("PASSO: Limpando campo de data inicial do filtro...")
        SELETOR_CAMPO_DATA_INICIAL_FILTRO = "#FiltroFaturamentoWrapper_dataInicial" # VERIFIQUE este ID!
        await page.locator(SELETOR_CAMPO_DATA_INICIAL_FILTRO).fill("") # Tenta preencher com string vazia

        print("INFO: Campo de data inicial limpo.")
        
        print("PASSO: Clicando em 'Aplicar' filtro...")
        SELETOR_BTN_APLICAR_FILTRO = "#aplicar" # VERIFIQUE! (Ou page.get_by_role("button", name="Aplicar"))
        await page.click(SELETOR_BTN_APLICAR_FILTRO)
        await page.wait_for_load_state('networkidle', timeout=45000) 
        print("INFO: Filtros aplicados.")

        # 3. Exportar Dados
        print("PASSO: Clicando em 'Exportar Dados' para Orçamento de Faturamento...")

        SELETOR_EXPORTAR_DADOS_ORCAMENTO = "#exportar_dados" 
        await page.click(SELETOR_EXPORTAR_DADOS_ORCAMENTO)
        
        # 3. Selecionar formato Excel/XLSX
        
        # >>> NOVA ESPERA ADICIONADA AQUI <<<
        print("INFO: Aguardando popup de exportação de clientes aparecer (pelo título)...")
        # Vamos esperar pelo título do popup que vemos na sua imagem

        print("PASSO: Selecionando formato Excel e XLSX no popup...")
        await page.locator("label").filter(has_text="EXCEL").nth(1).click()
        await page.locator("label").filter(has_text="xlsx").nth(1).click()
        await asyncio.sleep(2)
        
        print("PASSO: Clicando em 'Baixar' e aguardando download...")
        async with page.expect_download(timeout=60000) as download_info:
            await page.get_by_role("button", name=" EXPORTAR").click()
        
        print("INFO: Processo para baixar relatório de Orçamento de Faturamento concluído.")

        download = await download_info.value
        nome_arquivo_orcamento = f"Relatorio_ERP_Orc. Faturamento_OFICIAL_{id_filial_para_nome_arquivo}.xlsx" 
        caminho_pasta_downloads = hoje_formatado_local.replace('/', '-')
        if not os.path.exists(caminho_pasta_downloads):
            os.makedirs(caminho_pasta_downloads)
        caminho_para_salvar = f"{caminho_pasta_downloads}/{nome_arquivo_orcamento}"
        await download.save_as(caminho_para_salvar)
        print(f"INFO: Relatório de orçamento de faturamento para '{id_filial_para_nome_arquivo}' salvo como: {caminho_para_salvar}")
        print(f"INFO: Processo para baixar relatório de Orçamento de Faturamento para '{id_filial_para_nome_arquivo}' concluído.")
    except Exception as e_orc_fat:
        print(f"ERRO ao baixar relatório de Orçamento de Faturamento para '{id_filial_para_nome_arquivo}': {e_orc_fat}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção (Orçamento Faturamento - {id_filial_para_nome_arquivo})...")
            await page.pause()












async def baixar_pedidos_faturamento(page, datas, id_filial_para_nome_arquivo):
    print(f"\nINFO: Iniciando processo para baixar relatório de Pedidos de Faturamento para '{id_filial_para_nome_arquivo}'...")
    hoje_formatado_local = datas["hoje"]
    try:
        await abrir_menu_vendas(page) 
        print("PASSO: Clicando em 'Pedidos de faturamento'.")
        await page.click("#pedidofaturamentocrudcontroller")
        await page.wait_for_load_state('networkidle', timeout=30000)
        print("INFO: Página de Pedidos de Faturamento carregada.")
        await asyncio.sleep(1) 
        print("PASSO: Clicando no botão 'Filtrar' da grid...")
        await page.click("#filtrarGrid") 
        print("INFO: Aguardando popup de filtro...")
        await page.locator("#FiltroFaturamentoWrapper_status").wait_for(state="visible", timeout=15000)
        print("INFO: Popup de filtro aberto.")
        await asyncio.sleep(1)
        print("PASSO: Preenchendo 'Todos' no campo de status do filtro...")
        await page.locator("#FiltroFaturamentoWrapper_status").fill("Todos") 
        await asyncio.sleep(0.5)
        print("PASSO: Limpando campo de data inicial do filtro...")
        await page.locator("#FiltroFaturamentoWrapper_dataInicial").fill("") 
        print("PASSO: Clicando em 'Aplicar' filtro...")
        await page.click("#aplicar") 
        await page.wait_for_load_state('networkidle', timeout=45000) 
        print("INFO: Filtros aplicados para Pedidos.")
        print("PASSO: Clicando em 'Exportar Dados'...")
        await page.click("#exportar_dados") 
        print("INFO: Aguardando popup de exportação...")
        await page.locator("label").filter(has_text="EXCEL").nth(1).wait_for(state="visible", timeout=15000)
        print("PASSO: Selecionando formato Excel e XLSX no popup...")
        await page.locator("label").filter(has_text="EXCEL").nth(1).click() 
        await page.locator("label").filter(has_text="xlsx").nth(1).click() 
        await asyncio.sleep(2)
        print("PASSO: Clicando em 'Baixar' e aguardando download...")
        async with page.expect_download(timeout=60000) as download_info:
            await page.get_by_role("button", name=" EXPORTAR").click() 
        download = await download_info.value
        nome_arquivo_pedidos = f"Relatorio_ERP_Ped. Faturamento_OFICIAL_{id_filial_para_nome_arquivo}.xlsx" 
        caminho_pasta_downloads = hoje_formatado_local.replace('/', '-')
        if not os.path.exists(caminho_pasta_downloads):
            os.makedirs(caminho_pasta_downloads)
        caminho_para_salvar = f"{caminho_pasta_downloads}/{nome_arquivo_pedidos}"
        await download.save_as(caminho_para_salvar)
        print(f"INFO: Relatório de Pedidos de Faturamento para '{id_filial_para_nome_arquivo}' salvo como: {caminho_para_salvar}")
        print(f"INFO: Processo para baixar relatório de Pedidos de Faturamento para '{id_filial_para_nome_arquivo}' concluído.")
    except Exception as e_ped_fat:
        print(f"ERRO ao baixar relatório de Pedidos de Faturamento para '{id_filial_para_nome_arquivo}': {e_ped_fat}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção (Pedidos Faturamento - {id_filial_para_nome_arquivo})...")
            await page.pause()

async def baixar_nota_fiscal_entrada(page, datas, id_filial_para_nome_arquivo):
    print(f"\nINFO: Iniciando processo para baixar relatório de Nota Fiscal de Entrada para '{id_filial_para_nome_arquivo}'...")
    hoje_formatado_local = datas["hoje"] 
    try:
        await abrir_menu_fiscal(page)
        print("PASSO: Clicando em 'Nota Fiscal de Entrada'.")
        await page.click("#notafiscalentradacrudcontroller")
        await page.wait_for_load_state('networkidle', timeout=30000)
        await page.get_by_role("button", name="").wait_for(state="visible", timeout=15000)
        print("INFO: Página de Nota Fiscal de Entrada carregada.")
        print("PASSO: Clicando em 'Exportar Dados'...")
        await asyncio.sleep(1)
        await page.click("#exportar_dados")
        await asyncio.sleep(1)
        '''print('Vou clicar dnv')
        await page.click("#exportar_dados")
        print('cliquei dnv')'''
        
        print("INFO: Aguardando popup de exportação...")
        await page.locator("label").filter(has_text="EXCEL").nth(2).wait_for(state="visible", timeout=15000)
        print("INFO: Popup de exportação detectado.")
        print("PASSO: Selecionando formato Excel e XLSX no popup...")
        await page.locator("label").filter(has_text="EXCEL").nth(2).click() 
        await page.locator("label").filter(has_text="xlsx").nth(2).click() 
        await asyncio.sleep(1.5) 
        print("PASSO: Clicando em 'Baixar' DENTRO do popup e aguardando download...")
        async with page.expect_download(timeout=60000) as download_info:
            await page.get_by_role("button", name=" EXPORTAR").click()
        download = await download_info.value
        nome_arquivo_nf_entrada = f"Relatorio_ERP_Notas fiscais de entrada_OFICIAL_{id_filial_para_nome_arquivo}.xlsx" 
        caminho_pasta_downloads = hoje_formatado_local.replace('/', '-')
        if not os.path.exists(caminho_pasta_downloads):
            os.makedirs(caminho_pasta_downloads)
        caminho_para_salvar = f"{caminho_pasta_downloads}/{nome_arquivo_nf_entrada}"
        await download.save_as(caminho_para_salvar)
        print(f"INFO: Relatório de Nota Fiscal de Entrada para '{id_filial_para_nome_arquivo}' salvo como: {caminho_para_salvar}")
        print(f"INFO: Processo para baixar relatório de Nota Fiscal de Entrada para '{id_filial_para_nome_arquivo}' concluído.")
    except Exception as e_nf_entrada:
        print(f"ERRO ao baixar relatório de Nota Fiscal de Entrada para '{id_filial_para_nome_arquivo}': {e_nf_entrada}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção (NF Entrada - {id_filial_para_nome_arquivo})...")
            await page.pause()

async def baixar_nota_fiscal_saida(page, datas, id_filial_para_nome_arquivo):
    print(f"\nINFO: Iniciando processo para baixar relatório de Nota Fiscal de Saída para '{id_filial_para_nome_arquivo}'...")
    hoje_formatado_local = datas["hoje"] 
    try:
        await abrir_menu_fiscal(page) 
        print("PASSO: Clicando em 'Nota Fiscal de Saída'.")
        await page.click("#notafiscalsaidacrudcontroller") 
        await page.wait_for_load_state('networkidle', timeout=30000)
        print("INFO: Página de Nota Fiscal de Saída carregada.")
        print("PASSO: Clicando em 'Exportar Dados'...")
        await page.click("#exportar_dados") 
        print("INFO: Aguardando popup de exportação...")
        await page.locator("label").filter(has_text="EXCEL").nth(2).wait_for(state="visible", timeout=15000)
        print("INFO: Popup de exportação detectado.")
        print("PASSO: Selecionando formato Excel e XLSX no popup...")
        await page.locator("label").filter(has_text="EXCEL").nth(2).click() 
        await page.locator("label").filter(has_text="xlsx").nth(2).click() 
        await asyncio.sleep(1.5) 
        print("PASSO: Clicando em 'Baixar' DENTRO do popup e aguardando download...")
        async with page.expect_download(timeout=60000) as download_info:
            await page.get_by_role("button", name=" EXPORTAR").click()
        download = await download_info.value
        nome_arquivo_nf_saida = f"Relatorio_ERP_Notas fiscais de saída_OFICIAL_{id_filial_para_nome_arquivo}.xlsx" 
        caminho_pasta_downloads = hoje_formatado_local.replace('/', '-')
        if not os.path.exists(caminho_pasta_downloads):
            os.makedirs(caminho_pasta_downloads)
        caminho_para_salvar = f"{caminho_pasta_downloads}/{nome_arquivo_nf_saida}"
        await download.save_as(caminho_para_salvar)
        print(f"INFO: Relatório de Nota Fiscal de Saída para '{id_filial_para_nome_arquivo}' salvo como: {caminho_para_salvar}")
        print(f"INFO: Processo para baixar relatório de Nota Fiscal de Saída para '{id_filial_para_nome_arquivo}' concluído.")
    except Exception as e_nf_saida:
        print(f"ERRO ao baixar relatório de Nota Fiscal de Saída para '{id_filial_para_nome_arquivo}': {e_nf_saida}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção (NF Saída - {id_filial_para_nome_arquivo})...")
            await page.pause()











async def baixar_vendas_periodo_parametrizado(
    page, 
    datas, 
    id_filial_para_nome_arquivo, # Mantido para consistência, mas não usado no nome do arquivo aqui
    data_inicio_relatorio_str,    # Ex: datas_formatadas["doze_meses_atras"]
    seletor_opcao_agrupamento,    # O seletor Playwright para a opção de agrupamento
    texto_log_agrupamento,        # Descrição para o log, ex: "Mês" ou "Sem Agrupamento"
    seletor_formato_principal,    # Seletor para "EXCEL" ou "PDF" etc. (o botão maior)
    seletor_formato_tipo,         # Seletor para "xlsx" ou "xls" (a sub-opção)
    nome_arquivo_final_completo   # O nome completo do arquivo, ex: "vendasprodutosperiodo3meses.xls"
):
    print(f"\nINFO: Iniciando 'Vendas no Período' ({texto_log_agrupamento}, {nome_arquivo_final_completo}) para '{id_filial_para_nome_arquivo}'...")
    hoje_formatado_local = datas["hoje"]
    try:
        await abrir_menu_vendas(page)
        
        print("PASSO: Clicando em 'Vendas no Período'.")
        await page.click("#vendasprodutonoperiodoreportcontroller") 
        await page.wait_for_load_state('networkidle', timeout=45000)
        print("INFO: Página de 'Vendas no Período' carregada.")

        print("PASSO: Aplicando filtros...")
        print("PASSO: Abrindo dropdown de múltiplas filiais...")
        await asyncio.sleep(1)
        # Assumindo que os seletores de filial são os mesmos para todas as variações deste relatório
        await page.locator("it-filial-multiselect").get_by_role("button", name="Mostra lista").click()
        await asyncio.sleep(1)
        print("PASSO: Selecionando Filial 1 (checkbox)...")
        
        print("INFO: Filial 1 selecionada (checkbox).")
        await asyncio.sleep(1)
        print("PASSO: Selecionando Filial 2 (checkbox)...")
        await page.get_by_role("option", name="Nome: TRIMAF - TRIANGULO MATERIAIS GRAFICOS EIRELI Código: 2 CPF/CNPJ: 86.518.").get_by_label("").check()
        print("INFO: Filial 2 selecionada (checkbox).")
        await asyncio.sleep(1)
        await page.locator("it-filial-multiselect").get_by_role("button", name="Mostra lista").click() 
        print("INFO: Filiais selecionadas e dropdown fechado.")
        await asyncio.sleep(1)

        print(f"PASSO: Preenchendo data de início com: {data_inicio_relatorio_str}")
        await page.fill("#Report_inicio", data_inicio_relatorio_str)
        print("INFO: Data de início preenchida.")
        await asyncio.sleep(0.5)

        print(f"PASSO: Selecionando agrupamento: {texto_log_agrupamento}")
        await seletor_opcao_agrupamento.click() # Usa o seletor da opção de agrupamento
        print(f"INFO: Agrupamento '{texto_log_agrupamento}' selecionado.")
        await asyncio.sleep(0.5)

        print("PASSO: Selecionando formato de arquivo...")
        SELETOR_BASE_FORMATO_VP = "#footer-grid" 
        await page.locator(SELETOR_BASE_FORMATO_VP).locator(seletor_formato_principal).click() 
        await page.locator(SELETOR_BASE_FORMATO_VP).locator(seletor_formato_tipo).click()
        print(f"INFO: Formato de exportação selecionado.")
        await asyncio.sleep(1)

        print("PASSO: Clicando em 'Baixar/Agendar' e aguardando download...")
        timeout_download_vp_ms = 300 * 1000 
        async with page.expect_download(timeout=timeout_download_vp_ms) as download_info_vp:
           await page.click("#btn_rel_agendar") # Assumindo que este botão é o mesmo
        
        download_vp = await download_info_vp.value
        # Usando o nome de arquivo completo passado como parâmetro
        caminho_pasta_downloads_local = hoje_formatado_local.replace('/', '-')
        if not os.path.exists(caminho_pasta_downloads_local):
            os.makedirs(caminho_pasta_downloads_local)
        caminho_para_salvar_vp = f"{caminho_pasta_downloads_local}/{nome_arquivo_final_completo}"
        await download_vp.save_as(caminho_para_salvar_vp)
        print(f"INFO: Relatório '{nome_arquivo_final_completo}' salvo como: {caminho_para_salvar_vp}")
        print(f"INFO: Processo para baixar '{nome_arquivo_final_completo}' concluído.")
    except Exception as e_vp:
        print(f"ERRO ao baixar relatório '{nome_arquivo_final_completo}': {e_vp}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção ('{nome_arquivo_final_completo}')...")
            await page.pause()






async def baixar_posicao_estoque(page, datas):
    print("\nINFO: Iniciando processo para baixar relatório de 'Posição de Estoque'...")
    hoje_formatado_local = datas["hoje"]
    try:
        # 1. Navegação
        await abrir_menu_estoque(page)
        
        print("PASSO: Clicando em 'Posição de Estoque'.")
        SELETOR_LINK_POSICAO_ESTOQUE = "#posicaoestoquereportcontroller" # !! VERIFIQUE ESTE ID !!
        await page.click(SELETOR_LINK_POSICAO_ESTOQUE)
        await page.wait_for_load_state('networkidle', timeout=45000)
        print("INFO: Página de 'Posição de Estoque' carregada.")

        # 2. Selecionar Filiais (Múltipla Seleção)
        # Esta lógica será MUITO SIMILAR à seleção de filiais em "Giro de Estoque" e "Vendas no Período".
        # Você precisará encontrar o seletor do botão que abre o dropdown de filiais NESTA TELA.
        print("PASSO: Abrindo dropdown de múltiplas filiais para Posição de Estoque...")
        #await page.pause()
        await page.locator("it-empresa-filial").get_by_role("button", name="Mostra lista").click()
        await asyncio.sleep(1)

        print("PASSO: Selecionando Filial 1 (checkbox)...")
        # await page.pause()
        await page.get_by_role("option", name="Nome: TRIMAF - TRIANGULO MATERIAIS GRAFICOS EIRELI Código: 1 CPF/CNPJ: 86.518.").get_by_label("").check()

        print("PASSO: Selecionando Filial 2 (checkbox)...")
        # await page.pause()
        await page.get_by_role("option", name="Nome: TRIMAF - TRIANGULO MATERIAIS GRAFICOS EIRELI Código: 2 CPF/CNPJ: 86.518.").get_by_label("").check()
        
        # Como fechar o dropdown? (Ex: clicando no botão novamente, ou fora)
        # await page.click(SELETOR_PARA_ABRIR_DROPDOWN_FILIAL_POS_ESTOQUE) # Se clicar novamente fecha
        print("INFO: Filiais selecionadas. (Método para fechar dropdown a ser confirmado).")
        await asyncio.sleep(1)

        # 3. Selecionar Formato de Exportação (Excel, XLSX)
        print("PASSO: Selecionando formato Excel e XLSX...")
        
        # Use o Inspector para estes seletores DENTRO de SELETOR_BASE_FORMATO_VP 
        await page.locator("label").filter(has_text="EXCEL").first.first.click() 
        
        await page.locator("label").filter(has_text="xlsx").first.click()
        print("INFO: Formatos de exportação selecionados.")
        await asyncio.sleep(1)


        # 4. Clicar em Baixar
        print("PASSO: Clicando em 'Baixar/Agendar' e aguardando download...")
        # await page.pause()
        # Seu Selenium: //*[@id="btn_rel_agendar"]
        timeout_download_posicao_ms = 300 * 1000 # Ajuste se necessário
        async with page.expect_download(timeout=timeout_download_posicao_ms) as download_info_posicao:
           await page.get_by_role("button", name="b Gerar").click()
        
        download_posicao = await download_info_posicao.value
        nome_arquivo_posicao = f"Relatorio_ERP_Posicao de estoqueaAtual(uniplus)_OFICIAL.xlsx"
        caminho_pasta_downloads_local = hoje_formatado_local.replace('/', '-')
        
        if not os.path.exists(caminho_pasta_downloads_local):
            os.makedirs(caminho_pasta_downloads_local)

        caminho_para_salvar_posicao = f"{caminho_pasta_downloads_local}/{nome_arquivo_posicao}"
        await download_posicao.save_as(caminho_para_salvar_posicao)
        print(f"INFO: Relatório 'Posição de Estoque' salvo como: {caminho_para_salvar_posicao}")
        
        print("INFO: Processo para baixar 'Posição de Estoque' concluído.")

    except Exception as e_pos_est:
        print(f"ERRO ao baixar relatório de 'Posição de Estoque': {e_pos_est}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print("ERRO: Pausando para inspeção ('Posição de Estoque')...")
            await page.pause()








async def baixar_produtos_orcamento_novo(page, datas, id_filial_para_nome_arquivo):
    print(f"\nINFO: Iniciando processo para baixar relatório 'Produtos em Orçamento Novo' para '{id_filial_para_nome_arquivo}'...")
    hoje_formatado_local = datas["hoje"]
    dozemesesatras_local = datas["doze_meses_atras"]
    try:
        # 1. Navegação
        await abrir_menu_vendas(page)
        
        print("PASSO: Clicando em 'Produtos em Orçamento Novo'.")
        #await page.pause() # Descomente para encontrar este seletor
        await page.get_by_role("navigation").get_by_text("Produtos-orcamento-novo").click()
        await page.wait_for_load_state('networkidle', timeout=45000)
        print("INFO: Página de 'Produtos em Orçamento Novo' carregada.")

        # 2. Aplicar Filtros de Data
        print(f"PASSO: Preenchendo data inicial com: {dozemesesatras_local}")
        #await page.pause() # Use para encontrar o seletor do campo de data inicial
        SELETOR_DATA_INICIAL_PROD_ORC = "#ConfiguradorRelatorioParametros_datainicial" # << ENCONTRE ESTE!
        # Seu Selenium clicava e depois usava pyautogui.write. Com Playwright, tente .fill()
        await page.fill(SELETOR_DATA_INICIAL_PROD_ORC, dozemesesatras_local)
        print("INFO: Data inicial preenchida.")
        await asyncio.sleep(0.5)

        print(f"PASSO: Preenchendo data final com: {hoje_formatado_local}")
        # await page.pause() # Use para encontrar o seletor do campo de data final
        SELETOR_DATA_FINAL_PROD_ORC = "#ConfiguradorRelatorioParametros_datafinal" # << ENCONTRE ESTE!
        await page.fill(SELETOR_DATA_FINAL_PROD_ORC, hoje_formatado_local)
        print("INFO: Data final preenchida.")
        await asyncio.sleep(0.5)

        # 3. Selecionar Filial (se aplicável e se o controle for diferente)

        print("INFO: Selecionando Filial (checkbox)...")
        print("PASSO: Abrindo dropdown de Seleção de filial para produtos em Orçamento Novo...")
        #await page.pause()
        await page.locator("it-combobox").filter(has_text="Filial").get_by_label("Mostra lista").click()
        await asyncio.sleep(1)

        print("PASSO: Selecionando Filial 1 (checkbox)...")
        # await page.pause()
        await page.get_by_role("option", name="TRIMAF TRIANGULO MATERIAIS GRAFICOS EIRELI FILIAL 1").click()
        print("INFO: Filial 1 selecionada.")
        await asyncio.sleep(1)
        
        print("INFO: Filial selecionada. (Método para fechar dropdown a ser confirmado).")
        await asyncio.sleep(1)
        await page.locator("it-combobox").filter(has_text="Filial").get_by_label("Mostra lista").click()
        await asyncio.sleep(1)



        # 4. Selecionar Formato de Exportação (Apenas EXCEL no seu Selenium)
        print("PASSO: Selecionando formato EXCEL...")

        await page.locator("label").filter(has_text="EXCEL").first.first.click() 
        
        await page.locator("label").filter(has_text="xlsx").first.click()
        print("INFO: Formatos de exportação selecionados.")
        await asyncio.sleep(1)

        # 5. Clicar em Baixar/Gerar
        print("PASSO: Clicando em 'Baixar/Gerar' e aguardando download...")
        # O XPath do Selenium era: //*[@id="btn_rel_agendar"]/i (um ícone dentro do botão)
        # await page.pause() # Use para encontrar este seletor
        timeout_download_prod_orc_ms = 120 * 1000 # 2 minutos, ajuste se necessário
        
        async with page.expect_download(timeout=timeout_download_prod_orc_ms) as download_info_prod_orc:
           await page.get_by_role("button", name="b Gerar").click()
        
        download_prod_orc = await download_info_prod_orc.value
        nome_arquivo_prod_orc = f"Relatorio_ERP_RelatÃ³rio - Produtos-orcamento-novo (1)_OFICIAL.xlsx" 
        caminho_pasta_downloads_local = hoje_formatado_local.replace('/', '-')
        
        if not os.path.exists(caminho_pasta_downloads_local):
            os.makedirs(caminho_pasta_downloads_local)

        caminho_para_salvar_prod_orc = f"{caminho_pasta_downloads_local}/{nome_arquivo_prod_orc}"
        await download_prod_orc.save_as(caminho_para_salvar_prod_orc)
        print(f"INFO: Relatório 'Produtos em Orçamento Novo' para '{id_filial_para_nome_arquivo}' salvo como: {caminho_para_salvar_prod_orc}")
        
        print("INFO: Processo para baixar 'Produtos em Orçamento Novo' concluído.")

    except Exception as e_prod_orc:
        print(f"ERRO ao baixar relatório de 'Produtos em Orçamento Novo' para '{id_filial_para_nome_arquivo}': {e_prod_orc}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção ('Produtos em Orçamento Novo' - {id_filial_para_nome_arquivo})...")
            await page.pause()


            







async def baixar_produtos_orcamento_geral(page, datas, id_filial_para_nome_arquivo):
    print(f"\nINFO: Iniciando processo para baixar relatório 'Produtos em Orçamento (Geral)' para '{id_filial_para_nome_arquivo}'...")
    hoje_formatado_local = datas["hoje"]
    dozemesesatras_local = datas["doze_meses_atras"]
    try:
        # 1. Navegação
        await abrir_menu_vendas(page)
        
        print("PASSO: Clicando em 'Produtos em Orçamento (Geral)'.")

        await page.get_by_role("navigation").get_by_text("Produto em Orçamento").click()
        await page.wait_for_load_state('networkidle', timeout=45000)
        print("INFO: Página de 'Produtos em Orçamento (Geral)' carregada.")

        # 2. Aplicar Filtros de Data (ASSUMINDO QUE OS SELETORES DOS CAMPOS DE DATA SÃO OS MESMOS)
        print(f"PASSO: Preenchendo data inicial com: {dozemesesatras_local}")
        # await page.pause() # Use se o seletor de data for diferente ou precisar confirmar
        SELETOR_DATA_INICIAL_PROD_ORC_GERAL = "#ConfiguradorRelatorioParametros_datainicial" # << VERIFIQUE/CONFIRME (provavelmente o mesmo de "novo")
        await page.fill(SELETOR_DATA_INICIAL_PROD_ORC_GERAL, dozemesesatras_local)
        print("INFO: Data inicial preenchida.")
        await asyncio.sleep(0.5)

        print(f"PASSO: Preenchendo data final com: {hoje_formatado_local}")
        # await page.pause() 
        SELETOR_DATA_FINAL_PROD_ORC_GERAL = "#ConfiguradorRelatorioParametros_datafinal" # << VERIFIQUE/CONFIRME (provavelmente o mesmo de "novo")
        await page.fill(SELETOR_DATA_FINAL_PROD_ORC_GERAL, hoje_formatado_local)
        print("INFO: Data final preenchida.")
        await asyncio.sleep(0.5)

        # 3. Selecionar Filial (se aplicável e se o controle for diferente)

        print("INFO: Selecionando Filial (checkbox)...")
        print("PASSO: Abrindo dropdown de Seleção de filial para produtos em Orçamento Novo...")
        #await page.pause()
        await page.locator("it-combobox").filter(has_text="Filial").get_by_label("Mostra lista").click()
        await asyncio.sleep(1)

        print("PASSO: Selecionando Filial 1 (checkbox)...")
        # await page.pause()
        await page.get_by_role("option", name="TRIMAF TRIANGULO MATERIAIS GRAFICOS EIRELI FILIAL 1").click()
        print("INFO: Filial 1 selecionada.")
        await asyncio.sleep(1)
        
        print("INFO: Filial selecionada. (Método para fechar dropdown a ser confirmado).")
        await asyncio.sleep(1)
        await page.locator("it-combobox").filter(has_text="Filial").get_by_label("Mostra lista").click()
        await asyncio.sleep(1)



        # 4. Selecionar Formato de Exportação (Apenas EXCEL no seu Selenium)
        print("PASSO: Selecionando formato EXCEL...")

        await page.locator("label").filter(has_text="EXCEL").first.first.click() 
        
        await page.locator("label").filter(has_text="xlsx").first.click()
        print("INFO: Formatos de exportação selecionados.")
        await asyncio.sleep(1)

        # 5. Clicar em Baixar/Gerar
        print("PASSO: Clicando em 'Baixar/Gerar' e aguardando download...")
        # O XPath do Selenium era: //*[@id="btn_rel_agendar"]/i (um ícone dentro do botão)
        # await page.pause() # Use para encontrar este seletor
        timeout_download_prod_orc_ms = 120 * 1000 # 2 minutos, ajuste se necessário
        
        async with page.expect_download(timeout=timeout_download_prod_orc_ms) as download_info_prod_orc:
           await page.get_by_role("button", name="b Gerar").click()
        
        download_prod_orc = await download_info_prod_orc.value
        nome_arquivo_prod_orc = f"Relatorio_ERP_RelatÃ³rio - Produto em Orcamento 2.0_OFICIAL_{id_filial_para_nome_arquivo}.xlsx" # Nome do arquivo ajustado
        caminho_pasta_downloads_local = hoje_formatado_local.replace('/', '-')
        
        if not os.path.exists(caminho_pasta_downloads_local):
            os.makedirs(caminho_pasta_downloads_local)

        caminho_para_salvar_prod_orc = f"{caminho_pasta_downloads_local}/{nome_arquivo_prod_orc}"
        await download_prod_orc.save_as(caminho_para_salvar_prod_orc)
        print(f"INFO: Relatório 'Produtos em Orçamento (Geral)' para '{id_filial_para_nome_arquivo}' salvo como: {caminho_para_salvar_prod_orc}")
        
        print("INFO: Processo para baixar 'Produtos em Orçamento (Geral)' concluído.")

    except Exception as e_prod_orc_geral:
        print(f"ERRO ao baixar relatório de 'Produtos em Orçamento (Geral)' para '{id_filial_para_nome_arquivo}': {e_prod_orc_geral}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print(f"ERRO: Pausando para inspeção ('Produtos em Orçamento (Geral)' - {id_filial_para_nome_arquivo})...")
            await page.pause()



async def baixar_carteira_pedidos(page, datas, id_filial_para_nome_arquivo):
    print(f"\nINFO: Iniciando processo para baixar relatório 'Carteira de Pedidos' para '{id_filial_para_nome_arquivo}'...")
    hoje_formatado_local = datas["hoje"]
    
    try:
        # 1. Navegação
        # Assumindo que este relatório está no menu "Vendas". Se não estiver, mude para a função de menu correta.
        await abrir_menu_vendas(page)
        
        print("PASSO: Clicando em 'Carteira de Pedidos'.")
        # !! Use o Inspector para encontrar o seletor exato para o link/item 'Carteira de Pedidos' no menu. !!
        
        await page.get_by_text("Carteira de pedidos").click()
        await page.wait_for_load_state('networkidle', timeout=45000)
        print("INFO: Página de 'Carteira de Pedidos' carregada.")

        # 2. Selecionar Filiais (Múltipla Seleção)
        # Esta lógica será MUITO SIMILAR à seleção de filiais em "Giro de Estoque" e "Vendas no Período".
        # Você precisará encontrar o seletor do botão que abre o dropdown de filiais NESTA TELA.
        print("PASSO: Abrindo dropdown de múltiplas filiais para Posição de Estoque...")
        #await page.pause()
        await asyncio.sleep(1)
        await page.locator("it-empresa-filial").get_by_role("button", name="Mostra lista").click()
        await asyncio.sleep(1)

        print("PASSO: Selecionando Filial 1 (checkbox)...")
        # await page.pause()
        await page.get_by_role("option", name="Nome: TRIMAF - TRIANGULO MATERIAIS GRAFICOS EIRELI Código: 1 CPF/CNPJ: 86.518.").get_by_label("").check()

        await asyncio.sleep(1)

        print("PASSO: Selecionando Filial 2 (checkbox)...")
        # await page.pause()
        await page.get_by_role("option", name="Nome: TRIMAF - TRIANGULO MATERIAIS GRAFICOS EIRELI Código: 2 CPF/CNPJ: 86.518.").get_by_label("").check()
        
        # Como fechar o dropdown? (Ex: clicando no botão novamente, ou fora)
        # await page.click(SELETOR_PARA_ABRIR_DROPDOWN_FILIAL_POS_ESTOQUE) # Se clicar novamente fecha
        print("INFO: Filiais selecionadas. (Método para fechar dropdown a ser confirmado).")
        await asyncio.sleep(1)

        # 2. Aplicar Filtros Específicos
        print("PASSO: Aplicando filtros para 'Carteira de Pedidos'...")
        # Adicione um pause aqui para inspecionar todos os campos de filtro de uma vez
        # await page.pause()

        # 2a. Período: de '01/01/2020' até a data atual
        data_inicio_fixa = "01/01/2020"
        print(f"PASSO: Preenchendo data de início com: {data_inicio_fixa}")
        
        await page.locator("#Report_inicio").fill(data_inicio_fixa)
        
        print(f"PASSO: Preenchendo data final com: {hoje_formatado_local}")
        await page.locator("#Report_fim").fill(hoje_formatado_local)
        
        # 2b. Status -> todos
        print("PASSO: Preenchendo Status com 'Todos'.")
        await page.get_by_text("Todos").first.click()

        # 2c. "imprimir items" -> check
        print("PASSO: Marcando checkbox 'Imprimir Itens'.")
        # A melhor forma é usar get_by_label. Use o Inspector para confirmar o texto exato do label.
        await page.locator("it-checkbox").filter(has_text="Imprimir itens").locator("div").nth(2).check() # Use .check() para checkboxes
        
        # 2d. Estoque -> nenhum
        print("PASSO: Preenchendo Estoque com 'Nenhum'.")
        await page.get_by_text("Nenhum").click()

        # 2e. Impressão -> Somente Pedidos faturamento
        print("PASSO: Preenchendo Impressão com 'Somente Pedidos faturamento'.")
        await page.get_by_text("Somente pedidos de faturamento").click()




#------------------------------------------------
# 2f. Ordenar por -> "codigo"
        print("PASSO: Preenchendo Ordenar por com 'Código'.")
        SELETOR_ORDENAR_CARTEIRA = '//*[@id="indice"]/it-app/ng-component/ng-component/it-report-novo/it-component/div/form/fieldset/it-panel/div/div/div/it-row/div/it-panel[2]/div/div/div/it-row[7]/div/it-panel/div/div/div/it-radio/div/div/div[1]/fieldset/label/div' # << ENCONTRE ESTE!
        await page.click(SELETOR_ORDENAR_CARTEIRA) # Use o texto exato da opção

        # 2g. Agrupar por -> cliente
        print("PASSO: Preenchendo Agrupar por com 'Cliente'.")
        SELETOR_AGRUPAR_CARTEIRA = '//*[@id="indice"]/it-app/ng-component/ng-component/it-report-novo/it-component/div/form/fieldset/it-panel/div/div/div/it-row/div/it-panel[2]/div/div/div/it-row[8]/div/it-panel/div/div/div/it-radio/div/div/div[1]/fieldset/label/div/span' # << ENCONTRE ESTE!
        await page.click(SELETOR_AGRUPAR_CARTEIRA)
        print("INFO: Todos os filtros foram preenchidos.")
        await asyncio.sleep(2) # Pausa para ver os filtros aplicados
#------------------------------------------------


        # 3. Gerar e Baixar o Relatório
        # Esta parte é similar aos outros relatórios, mas provavelmente não há um popup de exportação,
        # o download deve iniciar diretamente após clicar em um botão "Gerar", "Imprimir" ou "Baixar".
        print("PASSO: Clicando em 'Gerar/Baixar' e aguardando download...")
        # await page.pause() # Para encontrar o seletor do botão final
        
        timeout_download_carteira_ms = 180 * 1000 # 3 minutos, ajuste se necessário
        
        async with page.expect_download(timeout=timeout_download_carteira_ms) as download_info:
           await page.get_by_role("button", name="b Gerar").click()
        
        download = await download_info.value
        nome_arquivo = f"Relatorio_ERP_RelatÃ³rio - Carteira de pedidos_OFICIAL.xlsx" # Assumindo XLSX
        caminho_pasta_downloads = hoje_formatado_local.replace('/', '-')
        
        if not os.path.exists(caminho_pasta_downloads):
            os.makedirs(caminho_pasta_downloads)
            
        caminho_para_salvar = f"{caminho_pasta_downloads}/{nome_arquivo}"
        await download.save_as(caminho_para_salvar)
        print(f"INFO: Relatório 'Carteira de Pedidos' salvo como: {caminho_para_salvar}")
        
        print("INFO: Processo para baixar 'Carteira de Pedidos' concluído.")

    except Exception as e_carteira:
        print(f"ERRO ao baixar relatório de 'Carteira de Pedidos': {e_carteira}")
        import traceback
        traceback.print_exc()
        if page and not page.is_closed():
            print("ERRO: Pausando para inspeção ('Carteira de Pedidos')...")
            await page.pause()
















## Função principal para executar a automação   
async def run_automation():
    if sys.platform == "win32" and sys.version_info >= (3, 8):
        try:
            current_policy = asyncio.get_event_loop_policy()
            if not isinstance(current_policy, asyncio.WindowsSelectorEventLoopPolicy):
                asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
                print("INFO: Política do loop de eventos asyncio definida para WindowsSelectorEventLoopPolicy.")
        except Exception as e:
            print(f"AVISO: Não foi possível definir a política do loop WindowsSelectorEventLoopPolicy: {e}")

    print("INFO: Iniciando a automação com Playwright...")
    async with async_playwright() as p:
        browser = None
        page = None 
        try:
            browser = await p.chromium.launch(headless=False, args=['--start-maximized'])
            # O new_context com no_viewport=True é bom. Se quiser downloads em pastas específicas por contexto:
            # downloads_path = os.path.join(os.getcwd(), "downloads_playwright") # Define um caminho base
            # context = await browser.new_context(no_viewport=True, accept_downloads=True, downloads_path=downloads_path)
            # Se não especificar downloads_path no contexto, o playwright baixa para uma pasta temporária
            # e o download.save_as() move para onde você especificar. Sua abordagem atual com save_as é boa.
            context = await browser.new_context(no_viewport=True, accept_downloads=True)
            page = await context.new_page()
            print("INFO: Navegador iniciado.")
            
            login_url_val = "https://compucenter-getcard03.getcard.uniplusweb.com/?p=endxQm15eXV4PzQ0aHRydXpoanN5ancybGp5aGZ3aTU4M2xqeWhmd2kzenNudXF6eHxqZzNodHIreWpzZnN5Qj07OjY9PTs8NTU1Njo%2B"
            usuario_val = "manoel"
            senha_val = "197322"

            await fazer_login(page, login_url_val, usuario_val, senha_val)
            
            # --- FILIAL 1 (TRIMAF) ---
            # Assegura que estamos na Filial 1 (TRIMAF) primeiro.
            # O nome da opção deve ser EXATAMENTE como no dropdown.
            NOME_FILIAL_1_DROPDOWN = "Nome: TRIMAF - TRIANGULO MATERIAIS GRAFICOS EIRELI Código: 1 CPF/CNPJ: 86.518."
            # Chamar selecionar_filial_padrao se ela já seleciona a TRIMAF, ou mudar_para_filial_especifica
            print(f"\n--- GARANTINDO/SELECIONANDO FILIAL 1 ({NOME_FILIAL_1_DROPDOWN}) ---")
            # Se selecionar_filial_padrao já faz isso, pode usá-la.
            # Ou, se você quiser ser explícito com a nova função:
            sucesso_filial_1 = await mudar_para_filial_especifica(page, NOME_FILIAL_1_DROPDOWN)
            # Alternativamente, se selecionar_filial_padrao já seleciona TRIMAF:
            # sucesso_filial_1 = await selecionar_filial_padrao(page)


            if sucesso_filial_1:
                print("\n--- INICIANDO DOWNLOADS PARA FILIAL 1 ---")
                id_f1 = "F1" # Identificador para os nomes dos arquivos
                #await asyncio.sleep(3)
                #await baixar_relatorio_giro_estoque(page, datas_formatadas)
                await baixar_carteira_pedidos(page, datas_formatadas, id_f1)
                await asyncio.sleep(3)
                await baixar_relatorio_produtos(page, datas_formatadas, id_f1) 
                
                await asyncio.sleep(3) 
                await baixar_relatorio_clientes(page, datas_formatadas)
                await asyncio.sleep(3) 
                await baixar_orcamento_faturamento(page, datas_formatadas, id_f1)
                await asyncio.sleep(3)
                await baixar_pedidos_faturamento(page, datas_formatadas, id_f1)
                await asyncio.sleep(3)
                await baixar_nota_fiscal_entrada(page, datas_formatadas, id_f1)
                await asyncio.sleep(3)
                await baixar_nota_fiscal_saida(page, datas_formatadas, id_f1)
                await asyncio.sleep(3)
                await baixar_posicao_estoque(page, datas_formatadas)
                await asyncio.sleep(3)
                await baixar_produtos_orcamento_novo(page, datas_formatadas, id_f1)
                await asyncio.sleep(3)
                await baixar_produtos_orcamento_geral(page, datas_formatadas, id_f1)
                await asyncio.sleep(3)
                print("\n--- INICIANDO VARIAÇÕES DO RELATÓRIO 'VENDAS NO PERÍODO' ---")

                # Variação 1: Ano Corrente (12 meses), Agrupamento por Mês, XLSX
                await baixar_vendas_periodo_parametrizado(
                    page, datas_formatadas, id_f1,
                    data_inicio_relatorio_str=datas_formatadas["doze_meses_atras"],
                    seletor_opcao_agrupamento=page.get_by_text("Mês"), # Seletor que você já usou
                    texto_log_agrupamento="Mês",
                    seletor_formato_principal=page.locator("label").filter(has_text="EXCEL").first, # Ou .nth(x)
                    seletor_formato_tipo=page.locator("label").filter(has_text="xlsx").first,    # Ou .nth(x)
                    nome_arquivo_final_completo="Relatorio_ERP_vendasprodutosperiodoAnoCorrente_OFICIAL.xlsx"
                )
                await asyncio.sleep(3)

                # Variação 2: Ano Corrente (12 meses), Sem Agrupamento, XLS
                # !! PRECISA ENCONTRAR O SELETOR PARA "SEM AGRUPAMENTO" !!
                # Ex: page.get_by_text("Nenhum") ou similar
                TEXTO_LOG_AGRUP_NENHUM = "Sem Agrupamento" # Ajuste o texto para o log
                # !! PRECISA ENCONTRAR/CONFIRMAR O SELETOR PARA A OPÇÃO "xls" !!
                SELETOR_TIPO_XLS = "label:has-text('xls')" # Exemplo, CONFIRME!

                await baixar_vendas_periodo_parametrizado(
                    page, datas_formatadas, id_f1,
                    data_inicio_relatorio_str=datas_formatadas["doze_meses_atras"], # Ano Corrente
                    seletor_opcao_agrupamento=page.get_by_text("Sem agrupamento"),
                    texto_log_agrupamento=TEXTO_LOG_AGRUP_NENHUM,
                    seletor_formato_principal=page.locator("label").filter(has_text="EXCEL").first, # Ou .nth(x)
                    seletor_formato_tipo=page.locator(SELETOR_TIPO_XLS).first, # Ou .nth(x)
                    nome_arquivo_final_completo="Relatorio_ERP_vendasprodutosperiodoAnoCorrente_OFICIAL.xls"
                )
                await asyncio.sleep(3)

                # Variação 3: 12 Meses (igual ao Ano Corrente), Agrupamento por Mês, XLS
                await baixar_vendas_periodo_parametrizado(
                    page, datas_formatadas, id_f1,
                    data_inicio_relatorio_str=datas_formatadas["doze_meses_atras"],
                    seletor_opcao_agrupamento=page.get_by_text("Mês"),
                    texto_log_agrupamento="Mês",
                    seletor_formato_principal=page.locator("label").filter(has_text="EXCEL").first,
                    seletor_formato_tipo=page.locator(SELETOR_TIPO_XLS).first, # Usando XLS
                    nome_arquivo_final_completo="Relatorio_ERP_vendasprodutosperiodo12meses_OFICIAL.xls"
                )
                await asyncio.sleep(3)

                # Variação 4: 6 Meses, Agrupamento por Mês, XLS
                await baixar_vendas_periodo_parametrizado(
                    page, datas_formatadas, id_f1,
                    data_inicio_relatorio_str=datas_formatadas["seis_meses_atras"],
                    seletor_opcao_agrupamento=page.get_by_text("Mês"),
                    texto_log_agrupamento="Mês",
                    seletor_formato_principal=page.locator("label").filter(has_text="EXCEL").first,
                    seletor_formato_tipo=page.locator(SELETOR_TIPO_XLS).first,
                    nome_arquivo_final_completo="Relatorio_ERP_vendasprodutosperiodo6meses_OFICIAL.xls"
                )
                await asyncio.sleep(3)

                # Variação 5: 3 Meses, Agrupamento por Mês, XLS
                await baixar_vendas_periodo_parametrizado(
                    page, datas_formatadas, id_f1,
                    data_inicio_relatorio_str=datas_formatadas["tres_meses_atras"],
                    seletor_opcao_agrupamento=page.get_by_text("Mês"),
                    texto_log_agrupamento="Mês",
                    seletor_formato_principal=page.locator("label").filter(has_text="EXCEL").first,
                    seletor_formato_tipo=page.locator(SELETOR_TIPO_XLS).first,
                    nome_arquivo_final_completo="Relatorio_ERP_vendasprodutosperiodo3meses_OFICIAL.xls"
                )
                await asyncio.sleep(3)

            else:
                print(f"AVISO: Não foi possível selecionar a Filial 1 ({NOME_FILIAL_1_DROPDOWN}). Downloads da Filial 1 abortados.")

            # --- MUDAR PARA FILIAL 2 E REPETIR DOWNLOADS ---
            print("\n--- TENTANDO MUDAR PARA FILIAL 2 ---")
            # !! IMPORTANTE: VOCÊ PRECISA DESCOBRIR O TEXTO EXATO DA OPÇÃO DA FILIAL 2 NO DROPDOWN !!
            # Use o Playwright Inspector na função selecionar_filial_padrao ou mudar_para_filial_especifica
            # para ver os nomes exatos das opções no dropdown de filiais.
            NOME_EXATO_FILIAL_2_DROPDOWN = "Nome: TRIMAF - TRIANGULO MATERIAIS GRAFICOS EIRELI Código: 2 CPF/CNPJ: 86.518." # << ENCONTRE E SUBSTITUA ISTO
            ID_ARQUIVO_FILIAL_2 = "F2" # Ex: filial2_XYZ (para nomear os arquivos)

            # Só tenta mudar para filial 2 e baixar se a filial 1 foi processada com sucesso (ou ajuste essa lógica)
            if sucesso_filial_1: # Ou remova esta condição se quiser tentar filial 2 independentemente
                sucesso_mudanca_filial2 = await mudar_para_filial_especifica(page, NOME_EXATO_FILIAL_2_DROPDOWN)

                if sucesso_mudanca_filial2:
                    print(f"\n--- INICIANDO DOWNLOADS PARA FILIAL 2 ({NOME_EXATO_FILIAL_2_DROPDOWN}) ---")
                    await asyncio.sleep(3)
                    # Relatórios que precisam ser baixados para a Filial 2 (baseado no seu notebook Selenium):
                    # Produtos, Orçamento, Pedidos, NF Entrada, NF Saída.
                    # Você pode adicionar/remover conforme necessário.
                    await baixar_relatorio_produtos(page, datas_formatadas, ID_ARQUIVO_FILIAL_2)
                    await asyncio.sleep(3)
                    await baixar_orcamento_faturamento(page, datas_formatadas, ID_ARQUIVO_FILIAL_2)
                    await asyncio.sleep(3)
                    await baixar_pedidos_faturamento(page, datas_formatadas, ID_ARQUIVO_FILIAL_2)
                    await asyncio.sleep(3)
                    await baixar_nota_fiscal_entrada(page, datas_formatadas, ID_ARQUIVO_FILIAL_2)
                    await asyncio.sleep(3)
                    await baixar_nota_fiscal_saida(page, datas_formatadas, ID_ARQUIVO_FILIAL_2)
                    # Adicione outros se necessário (Giro, Vendas Período, Posição Estoque para Filial 2?)
                    await asyncio.sleep(3)
                    await baixar_relatorio_giro_estoque(page, datas_formatadas)
                else:
                    print(f"AVISO: Não foi possível mudar para a Filial 2. Downloads para Filial 2 abortados.")
            
            # --- PRÓXIMOS RELATÓRIOS DISTINTOS (APÓS LÓGICA DE FILIAIS) ---
            # Ex: Produtos em orçamento (Células 24 e 25 do seu notebook Selenium)
            # Você criará novas funções para eles, como por exemplo:
            # await asyncio.sleep(3)
            # await baixar_produtos_orcamento_novo(page, datas_formatadas, "filial_ativa_ou_geral") 
            # await asyncio.sleep(3)
            # await baixar_produtos_orcamento_antigo(page, datas_formatadas, "filial_ativa_ou_geral")

            print("\nINFO: Todas as tarefas de download programadas foram tentadas.")

        except Exception as e_playwright:
            print(f"ERRO GERAL NA AUTOMACAO: {e_playwright}")
            import traceback
            traceback.print_exc()
            if page and not page.is_closed():
                print("ERRO: Pausando para inspeção devido a erro geral...")
                await page.pause()
        finally:
            if browser:
                print("INFO: Fechando o navegador...")
                await browser.close()
            print("INFO: Finalizando a função de automação.")


if __name__ == "__main__":
    print("INFO: Executando o script de automação 'automacao_playwright.py'...")
    asyncio.run(run_automation())
    print("INFO: Script finalizado.")