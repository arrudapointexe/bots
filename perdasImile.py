import os
import time
import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

def rodar_extracao_vencimentos():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(no_viewport=True, accept_downloads=True)
        page = context.new_page()

        print("1. Acessando o portal e fazendo login...")
        page.goto("https://ds-login.imile.com/", wait_until="domcontentloaded")
        page.fill('input[type="text"]', os.getenv("IMILE_USER"))
        page.fill('input[type="password"]', os.getenv("IMILE_PASS"))
        page.get_by_role("button", name="登录").click()
        
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(5000)

        # Limpa pop-ups
        btn_close = page.locator('.close-icon').first
        btn_close.click(force=True)

        print("2. Navegando para Monitor de Vencimento...")
        page.get_by_text("Monitoramento").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Monitor de Operação").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Monitor de vencimento").first.click(force=True)
        
        print("Aguardando tabela carregar...")
        page.wait_for_timeout(8000)

        # 3. Clica na primeira linha, segunda coluna (Onde fica a numeração azul)
        print("Abrindo modal de vencimentos...")
        celula_alvo = page.locator('.imile-table-row-level-0').first.locator('td').nth(1)
        celula_alvo.click(force=True)
        
        page.wait_for_timeout(3000) # Aguarda o modal abrir

        print("4. Extraindo AWBs de todas as páginas...")
        codigos = []
        
        while True:
            # Puxa todos os códigos da página atual baseado no HTML que você enviou
            elementos = page.locator('tbody.imile-table-tbody tr.imile-table-row').all()
            for el in elementos:
                try:
                    # Pega o atributo data-row-key que contém o AWB (ex: 6022826682795)
                    awb = el.get_attribute("data-row-key")
                    if awb and awb.strip():
                        codigos.append(awb.strip())
                except: continue
            
            print(f"Total capturado até agora: {len(codigos)}")
            
            # --- TENTA IR PARA A PRÓXIMA PÁGINA ---
            # ATENÇÃO: Ajuste o seletor abaixo de acordo com a setinha de passar página da iMile
            btn_proxima = page.locator('.MuiPaginationItem-next, .ant-pagination-next, button[aria-label="Go to next page"]').first
            
            if btn_proxima.is_visible() and not btn_proxima.is_disabled():
                btn_proxima.click(force=True)
                page.wait_for_timeout(2000) # Pausa para carregar a nova lista
            else:
                break # Chegou na última página
                
        codigos = list(set(codigos)) # Remove duplicados por segurança
        print(f"Captura finalizada! Total de AWBs únicos: {len(codigos)}")
        
        if not codigos:
            print("Nenhum código encontrado.")
            browser.close()
            return

        # 5. Navega para a aba Controle de Pedidos
        print("5. Indo para a tela de Pesquisa...")
        # NOTA: Ajuste o nome do menu se for diferente no seu portal
        page.get_by_text("Controle de Pedidos", exact=False).first.click(force=True) 
        page.wait_for_timeout(5000)

        # Muda o tipo de pesquisa (Rastreamento -> Pesquisar por destinatário)
        print("Colando códigos...")
        try:
            page.get_by_text("Rastreamento").first.click(force=True)
            page.get_by_text("Pesquisar por destinatário").first.click(force=True)
            page.wait_for_timeout(1000)
        except: pass

        # Clica no input e cola os dados separando por vírgula ou quebra de linha
        campo_busca = page.locator('input[placeholder="Por favor, insira"]').first
        campo_busca.click()
        campo_busca.fill("\n".join(codigos)) # O sistema da iMile costuma aceitar múltiplos com quebra de linha
        page.wait_for_timeout(1000)

        # 6. Processo de Exportação
        print("6. Solicitando Extração...")
        # Clica no botão Extrair
        page.locator('.ImileActionButton-root:has-text("extrair")').first.click(force=True)
        page.wait_for_timeout(1000)
        
        # Clica em Exportar tudo
        page.locator('li.ImileMenuItem-root:has-text("Exportar tudo")').first.click(force=True)
        page.wait_for_timeout(1000)
        
        # Clica em Criar arquivo de exportação
        page.locator('.export-button').first.click(force=True)
        
        print("Aguardando a iMile gerar o arquivo (isso pode demorar)...")
        page.wait_for_timeout(10000) # Tempo base para a iMile processar
        
        # 7. Baixando o arquivo
        print("7. Baixando o arquivo...")
        # Usa o expect_download para interceptar o download pelo Playwright
        with page.expect_download(timeout=60000) as download_info:
            # Clica no ícone da nuvem de download que você mandou o HTML
            page.locator('span.Imile-ButtonIcon-root svg path[d*="M8 0C3.57"]').locator('..').first.click(force=True)
            
        download = download_info.value
        caminho_arquivo = os.path.join(os.getcwd(), "base_vencimentos.xlsx")
        download.save_as(caminho_arquivo)
        print(f"Arquivo salvo com sucesso em: {caminho_arquivo}")
        
        browser.close()
        
        # 8. Chama a função Pandas para gerar a tabela
        gerar_relatorio_pandas(caminho_arquivo)


def gerar_relatorio_pandas(caminho_arquivo):
    print("\n" + "="*50)
    print("📊 GERANDO RELATÓRIO DINÂMICO (PANDAS)")
    print("="*50)
    
    try:
        # Lê a planilha exportada
        df = pd.read_excel(caminho_arquivo)
        
        # ATENÇÃO: Você precisa checar os nomes exatos das colunas no arquivo Excel que a iMile gera.
        # Estou supondo que se chamam "Motorista" e "Status". Altere abaixo se for "Nome do Motorista", etc.
        coluna_motorista = "Motorista" 
        coluna_status = "Status"
        
        # Cria a tabela cruzada (Pivot Table) contando a quantidade de cada status por motorista
        tabela = pd.crosstab(
            df[coluna_motorista], 
            df[coluna_status], 
            margins=True,          # Adiciona a linha/coluna de "Total Geral"
            margins_name="Total Geral"
        )
        
        # Converte a tabela pra string para facilitar copiar para o WhatsApp
        relatorio_texto = tabela.to_string()
        
        print(relatorio_texto)
        
        # Salva o relatório num .txt para você só copiar e colar
        with open("relatorio_vencimentos_pronto.txt", "w", encoding="utf-8") as f:
            f.write("VENCIMENTO IMILE 06/04\n")
            f.write("="*50 + "\n")
            f.write(relatorio_texto)
            f.write("\n" + "="*50 + "\n")
            
        print("\n✅ Relatório gerado e salvo em 'relatorio_vencimentos_pronto.txt'!")
        
    except Exception as e:
        print(f"Erro ao gerar relatório Pandas: {e}")
        print("DICA: Abra o arquivo base_vencimentos.xlsx, veja o nome exato das colunas de Motorista e Status e atualize no código.")

if __name__ == "__main__":
    rodar_extracao_vencimentos()