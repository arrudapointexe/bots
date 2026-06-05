import os
import re
from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from imile_utils import login_imile
from google_utils import get_gspread_client

load_dotenv()

NOME_PLANILHA = os.getenv("NOME_PLANILHA", "acareaBase")

def salvar_no_historico(sigla, entregues, acareacoes):
    """Salva os dados consolidados do dia na aba KPI_Historico"""
    print(f"[{sigla}] Salvando resultados no Google Sheets...")
    try:
        cliente = get_gspread_client()
        planilha_mestre = cliente.open(NOME_PLANILHA)
        
        nome_aba_kpi = "KPI_Historico"
        try:
            planilha_kpi = planilha_mestre.worksheet(nome_aba_kpi)
        except gspread.exceptions.WorksheetNotFound:
            planilha_kpi = planilha_mestre.add_worksheet(title=nome_aba_kpi, rows="1000", cols="6")
            planilha_kpi.append_row(["Data", "Base", "Entregues", "Acareações", "KPI (%)"])

        # Calcula o KPI (Porcentagem de Acareações sobre os Entregues)
        kpi_perc = 0.0
        if entregues > 0:
            kpi_perc = round((acareacoes / entregues) * 100, 2)
            
        # Pega a data de ontem (já que o Timeliness Report do dia anterior é o que está consolidado)
        data_ontem = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

        planilha_kpi.append_row([data_ontem, sigla, entregues, acareacoes, f"{kpi_perc}%"])
        print(f"[{sigla}] ✅ KPI salvo com sucesso: {kpi_perc}%")
        
    except Exception as e:
        print(f"❌ Erro ao salvar KPI da base {sigla}: {e}")


def extrair_entregas_imile(page, sigla):
    """Navega até o Timeliness Report, preenche a data e clica em Pesquisar"""
    try:
        print(f"[{sigla}] 1/2: Extraindo Entregas (Timeliness Report)...")
        page.get_by_text("Monitor").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Operation Monitor").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Timeliness Report").first.click(force=True)
        page.wait_for_timeout(8000)

        # ==========================================
        # FILTRAR PELA DATA DE ONTEM E PESQUISAR
        # ==========================================
        data_ontem = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        print(f"[{sigla}] -> Aplicando filtro de data forçado (Entregas): {data_ontem}")
        
        for frame in page.frames:
            try:
                inputs_data = frame.locator('input[placeholder*="ate"], input[placeholder*="ata"], input[placeholder*="Start"], input[placeholder*="End"], input[placeholder*="Time"]').all()
                datas_encontradas = [inp for inp in inputs_data if inp.is_visible()]
                
                if len(datas_encontradas) >= 2:
                    for i in range(2):
                        datas_encontradas[i].click(force=True)
                        page.wait_for_timeout(500)
                        datas_encontradas[i].press("Control+A")
                        datas_encontradas[i].press("Backspace")
                        datas_encontradas[i].type(data_ontem, delay=150)
                        datas_encontradas[i].press("Enter")
                        page.wait_for_timeout(500)
                    
                    # 👇 CORREÇÃO: Aperta "ESC" no teclado para fechar o calendário com segurança
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(1000)
                    
                    # Procura o botão de Pesquisar ampliando os alvos
                    btn_search = frame.locator('.search-btn, button:has-text("Search"), button:has-text("Pesquisar"), button:has-text("Query"), .ant-btn-primary').first
                    
                    if btn_search.is_visible(timeout=3000):
                        btn_search.click(force=True)
                        print(f"[{sigla}] -> Pesquisa clicada! Aguardando o servidor carregar as entregas...")
                    else:
                        # PLANO B: Se o botão estiver escondido, força o clique via JavaScript
                        frame.evaluate('''() => {
                            let btns = Array.from(document.querySelectorAll('button'));
                            let searchBtn = btns.find(b => b.innerText.includes('Search') || b.innerText.includes('Pesquisar') || b.innerText.includes('Query'));
                            if(searchBtn) searchBtn.click();
                        }''')
                        print(f"[{sigla}] -> Pesquisa forçada via JS! Aguardando o servidor...")
                        
                    page.wait_for_timeout(12000)
                    break
            except:
                continue

        # ==========================================
        # LER O VALOR DA TABELA
        # ==========================================
        entregues_str = page.evaluate('''() => {
            const headers = Array.from(document.querySelectorAll("thead th")).map(h => h.innerText.trim());
            const index = headers.findIndex(h => h.includes("On-time Delivery Quantity") || h.includes("On-time"));
            
            if (index >= 0) {
                const firstRow = document.querySelector(".imile-table-row-level-0");
                if (firstRow) {
                    const cells = firstRow.querySelectorAll("td");
                    if (cells && cells[index]) {
                        return cells[index].innerText.trim();
                    }
                }
            }
            return "0";
        }''')
        
        qtd_entregues = int(entregues_str.replace(',', '').replace('.', '')) if entregues_str else 0
        print(f"[{sigla}] -> Pacotes entregues (On-time): {qtd_entregues}")
        return qtd_entregues

    except Exception as e:
        print(f"[{sigla}] ❌ Erro ao extrair Timeliness Report: {e}")
        return 0


def extrair_acareacoes_imile(page, sigla):
    """Navega até o Complaint Ticket, lê a coluna Create Time e conta as acareações geradas ONTEM"""
    try:
        print(f"[{sigla}] 2/2: Lendo a tabela para capturar a data de criação...")
        page.get_by_text("Service").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Ticket management").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Complaint ticket").first.click(force=True)
        page.wait_for_timeout(10000)

        # ==========================================
        # 1. AJUSTAR PAGINAÇÃO PARA 100
        # ==========================================
        print(f"[{sigla}] -> Ajustando paginação para 100...")
        for frame in page.frames:
            try:
                caixa_paginacao = frame.locator('.imile-pagination-options .imile-select-selector').first
                if caixa_paginacao.is_visible(timeout=1000):
                    texto_atual = frame.locator('.imile-pagination-options .imile-select-selection-item').first.inner_text()
                    if "100" not in texto_atual:
                        caixa_paginacao.click(force=True)
                        page.wait_for_timeout(1500)
                        opcao_100 = frame.locator('div[title="100 / página"], div[title="100 / page"]').first
                        if opcao_100.is_visible(timeout=2000):
                            opcao_100.click(force=True)
                            print(f"[{sigla}] -> Paginação alterada! Esperando 10 segundos para a tabela recarregar...")
                            page.wait_for_timeout(10000) # Tempo extra para a JML conseguir carregar
                    break 
            except: continue

        # ==========================================
        # 2. CAPTURAR E CONTAR AS DATAS DE ONTEM
        # ==========================================
        data_ontem = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        print(f"[{sigla}] -> Inspecionando coluna de data buscando: {data_ontem}...")

        total_ontem = 0
        continua_buscando = True
        paginas_lidas = 0
        max_paginas = 20 

        while continua_buscando and paginas_lidas < max_paginas:
            paginas_lidas += 1
            encontrou_tabela_na_pagina = False

            for frame in page.frames:
                try:
                    resultado = frame.evaluate('''() => {
                        const headers = Array.from(document.querySelectorAll("thead th")).map(h => h.innerText.trim().toLowerCase());
                        // Busca a coluna tolerando inglês e português
                        const idx = headers.findIndex(h => h.includes("create") || h.includes("criaç") || h.includes("criac"));
                        
                        if (idx === -1) return { datas: [], erro: "Sem coluna de data" };

                        const rows = Array.from(document.querySelectorAll("tr.imile-table-row"));
                        const datas = rows.map(row => {
                            const cells = row.querySelectorAll("td");
                            return cells[idx] ? cells[idx].innerText.trim() : "";
                        }).filter(d => d !== "");

                        return { datas: datas, erro: null };
                    }''')

                    if resultado and not resultado.get("erro"):
                        encontrou_tabela_na_pagina = True
                        datas_pagina = resultado["datas"]

                        for data_texto in datas_pagina:
                            if data_texto.startswith(data_ontem):
                                total_ontem += 1
                            elif data_texto < data_ontem:
                                continua_buscando = False
                                break

                        if len(datas_pagina) < 100:
                            continua_buscando = False

                        if continua_buscando:
                            btn_next = frame.locator('.imile-pagination-next, li[title="Next Page"], li[title="Próxima Página"]').first
                            if btn_next.is_visible():
                                classe_btn = btn_next.get_attribute("class") or ""
                                aria_disabled = btn_next.get_attribute("aria-disabled") or "false"
                                
                                if "disabled" in classe_btn or aria_disabled == "true":
                                    continua_buscando = False
                                else:
                                    print(f"[{sigla}] -> Lendo página {paginas_lidas}... Avançando para a próxima!")
                                    btn_next.click(force=True)
                                    page.wait_for_timeout(5000)
                            else:
                                continua_buscando = False
                        
                        break
                except:
                    continue
            
            if not encontrou_tabela_na_pagina:
                print(f"[{sigla}] ⚠️ Tabela não encontrada nesta página. Encerrando a busca...")
                continua_buscando = False

        print(f"[{sigla}] -> Contagem cirúrgica concluída! Acareações de {data_ontem}: {total_ontem}")
        return total_ontem

    except Exception as e:
        print(f"[{sigla}] ❌ Erro ao ler a tabela do Complaint Ticket: {e}")
        return 0


def rodar_kpi_base(sigla, usuario, senha, p):
    print(f"\n{'='*50}\n📊 CALCULANDO KPI: {sigla}\n{'='*50}")
    if not usuario or not senha:
        print(f"⚠️ Credenciais de {sigla} ausentes. Pulando...")
        return

    browser = p.chromium.launch(headless=False, args=["--start-maximized"])
    context = browser.new_context(no_viewport=True)
    page = context.new_page()

    try:
        # LOGIN
        login_imile(page, usuario, senha)

        # 1. Pega os Entregues
        entregues = extrair_entregas_imile(page, sigla)
        
        # 2. Pega as Acareações
        # (Para o robô não se perder, voltamos para a home antes de ir pro próximo menu)
        page.goto("https://ds.imile.com/#/dashboard", wait_until="domcontentloaded")
        page.wait_for_timeout(3000)
        acareacoes = extrair_acareacoes_imile(page, sigla)
        
        # 3. Salva no Google Sheets
        salvar_no_historico(sigla, entregues, acareacoes)

    except Exception as e:
        print(f"❌ Erro crítico no KPI da base {sigla}: {e}")
    finally:
        browser.close()


def main():
    bases = [
        ("CTG", os.getenv("IMILE_CTG_USER"), os.getenv("IMILE_CTG_PASS")),
        ("ITR", os.getenv("IMILE_ITR_USER"), os.getenv("IMILE_ITR_PASS")),
        ("JML", os.getenv("IMILE_JML_USER"), os.getenv("IMILE_JML_PASS"))
    ]

    with sync_playwright() as p:
        for sigla, usuario, senha in bases:
            rodar_kpi_base(sigla, usuario, senha, p)

if __name__ == "__main__":
    main()