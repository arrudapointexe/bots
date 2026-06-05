import os
import re
from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import gspread
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv()

NOME_PLANILHA = os.getenv("NOME_PLANILHA", "acareaBase")

def salvar_no_historico(sigla, entregues, acareacoes, data_alvo):
    """Salva os dados retroativos na aba KPI_Historico"""
    try:
        escopo = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credenciais = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', escopo)
        cliente = gspread.authorize(credenciais)
        planilha_mestre = cliente.open(NOME_PLANILHA)
        
        nome_aba_kpi = "KPI_Historico"
        try:
            planilha_kpi = planilha_mestre.worksheet(nome_aba_kpi)
        except gspread.exceptions.WorksheetNotFound:
            planilha_kpi = planilha_mestre.add_worksheet(title=nome_aba_kpi, rows="1000", cols="6")
            planilha_kpi.append_row(["Data", "Base", "Entregues", "Acareações", "KPI (%)"])

        kpi_perc = 0.0
        if entregues > 0:
            kpi_perc = round((acareacoes / entregues) * 100, 2)
            
        planilha_kpi.append_row([data_alvo, sigla, entregues, acareacoes, f"{kpi_perc}%"])
        print(f"[{sigla} - {data_alvo}] ✅ Salvo no Sheets: Entregues: {entregues} | Acareações: {acareacoes} | KPI: {kpi_perc}%")
        
    except Exception as e:
        print(f"❌ Erro ao salvar KPI retroativo da base {sigla} ({data_alvo}): {e}")

def extrair_entregas_imile(page, sigla, data_alvo):
    try:
        print(f"[{sigla} - {data_alvo}] 1/2: Extraindo Entregas (Timeliness Report)...")
        page.goto("https://ds.imile.com/#/dashboard", wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        
        page.get_by_text("Monitor").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Operation Monitor").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Timeliness Report").first.click(force=True)
        page.wait_for_timeout(6000)

        dia_alvo = str(int(data_alvo.split('-')[2]))
        
        # 1. ABRE O CALENDÁRIO CLICANDO NO COMPONENTE VISUAL
        for frame in [page] + page.frames:
            try:
                trigger = frame.locator('.RangeDatePicker-root, .ImileEllipsis-root').first
                if trigger.is_visible(timeout=1000):
                    trigger.click(force=True)
                    page.wait_for_timeout(2000) # Espera abrir o pop-up
                    break
            except: continue

        # 2. SELEÇÃO DO DIA DENTRO DO POPOVER (MÊS ATUAL)
        for frame in [page] + page.frames:
            try:
                # Localiza a janela flutuante
                dialog = frame.locator('.z-imd-popover, div[role="tooltip"], div[role="dialog"]').last
                if not dialog.is_visible():
                    dialog = page.locator('.z-imd-popover, div[role="tooltip"], div[role="dialog"]').last
                
                if dialog.is_visible():
                    # Executa o duplo clique usando injeção limpa no primeiro painel do mês atual
                    clicou = dialog.evaluate(f'''(node) => {{
                        const diaProcurado = "{dia_alvo}";
                        // Pega os painéis de meses na tela (o primeiro é o mês atual)
                        const paineis = Array.from(node.querySelectorAll('.StaticDatePicker-root, .imd-panel-date'));
                        if (!paineis.length) return false;
                        
                        // Busca apenas os dias ativos (ignora prevMonth e nextMonth)
                        const celulas = Array.from(paineis[0].querySelectorAll('.StaticDayPicker-item:not(.prevMonth):not(.nextMonth), .imd-picker-cell-in-view'));
                        const alvo = celulas.find(c => c.innerText.replace(/\\D/g, '') === diaProcurado);
                        
                        if (alvo) {{
                            const btn = alvo.querySelector('.StaticDayPicker-innerItem, .imd-picker-date-value') || alvo;
                            btn.click(); // Clique 1 (Início)
                            setTimeout(() => btn.click(), 300); // Clique 2 (Fim) após mini-pausa
                            return true;
                        }}
                        return false;
                    }}''')
                    
                    if clicou:
                        print(f"[{sigla}] -> Duplo clique executado no dia {dia_alvo} com sucesso.")
                        page.wait_for_timeout(1500)
                    
                    # Clica em Confirmar no rodapé do calendário
                    btn_confirm = dialog.locator('.imd-picker-btn-confirm, button:has-text("Confirm"), button:has-text("Confirmar")').first
                    if btn_confirm.is_visible():
                        btn_confirm.click(force=True)
                        page.wait_for_timeout(1000)
                    else:
                        page.keyboard.press("Escape")
                    break
            except: continue

        # 3. CLICA NO BOTÃO SEARCH IMEDIATAMENTE
        page.wait_for_timeout(1000)
        pesquisou = False
        for frame in [page] + page.frames:
            try:
                btn_search = frame.locator('.search-btn, button:has-text("Search"), button:has-text("Pesquisar"), .ant-btn-primary').first
                if btn_search.is_visible(timeout=2000):
                    btn_search.click(force=True)
                    pesquisou = True
                    print(f"[{sigla}] -> Botão SEARCH acionado!")
                    break
            except: continue
            
        if not pesquisou:
            page.evaluate('''() => {
                let btns = Array.from(document.querySelectorAll('button'));
                let searchBtn = btns.find(b => b.innerText.includes('Search') || b.innerText.includes('Pesquisar'));
                if(searchBtn) searchBtn.click();
            }''')
        
        page.wait_for_timeout(10000)

        # 4. LEITURA DO VALOR DA TABELA
        entregues_str = page.evaluate('''() => {
            const headers = Array.from(document.querySelectorAll("thead th")).map(h => h.innerText.trim());
            const index = headers.findIndex(h => h.includes("On-time Delivery Quantity") || h.includes("On-time"));
            if (index >= 0) {
                const firstRow = document.querySelector(".imile-table-row-level-0");
                if (firstRow) {
                    const cells = firstRow.querySelectorAll("td");
                    if (cells && cells[index]) return cells[index].innerText.trim();
                }
            }
            return "0";
        }''')
        return int(entregues_str.replace(',', '').replace('.', '')) if entregues_str else 0
        
    except Exception as e:
        print(f"[{sigla}] ❌ Erro ao extrair Entregas: {e}")
        return 0

def extrair_acareacoes_imile(page, sigla, data_alvo):
    try:
        print(f"[{sigla} - {data_alvo}] 2/2: Extraindo Acareações (Complaint Ticket)...")
        page.goto("https://ds.imile.com/#/dashboard", wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        
        page.get_by_text("Service").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Ticket management").first.click(force=True)
        page.wait_for_timeout(1000)
        page.get_by_text("Complaint ticket").first.click(force=True)
        page.wait_for_timeout(8000)

        dia_alvo = str(int(data_alvo.split('-')[2]))

        # 1. ABRE O CALENDÁRIO CLICANDO NO INPUT INTERNO DO "CREATE TIME"
        # (Usando o HTML exato que você mandou da caixinha de texto)
        for frame in [page] + page.frames:
            try:
                # Alveja a barra do mistrigger ou o próprio input de texto que guarda a data
                trigger_input = frame.locator('[data-slot="mixtrigger-value"] input, .imd-pro-date-trigger input').first
                if not trigger_input.is_visible(timeout=1000):
                    trigger_input = frame.locator('input[value*="2026"]').first

                if trigger_input.is_visible(timeout=1000):
                    trigger_input.click(force=True)
                    print(f"[{sigla}] -> Clique no campo 'Create Time' executado. Abrindo pop-up...")
                    page.wait_for_timeout(2000)
                    break
            except: continue

        # 2. SELEÇÃO DO DIA DENTRO DO POPOVER (MÊS ATUAL)
        for frame in [page] + page.frames:
            try:
                dialog = frame.locator('.z-imd-popover, [data-ui-slot="popover-content"], div[role="dialog"]').last
                if not dialog.is_visible():
                    dialog = page.locator('.z-imd-popover, [data-ui-slot="popover-content"], div[role="dialog"]').last
                
                if dialog.is_visible():
                    clicou = dialog.evaluate(f'''(node) => {{
                        const diaProcurado = "{dia_alvo}";
                        const paineis = Array.from(node.querySelectorAll('.StaticDatePicker-root, .imd-panel-date, .imd-picker-range'));
                        if (!paineis.length) return false;
                        
                        const celulas = Array.from(paineis[0].querySelectorAll('.StaticDayPicker-item:not(.prevMonth):not(.nextMonth), .imd-picker-cell-in-view'));
                        const alvo = celulas.find(c => c.innerText.replace(/\\D/g, '') === diaProcurado);
                        
                        if (alvo) {{
                            const btn = alvo.querySelector('.StaticDayPicker-innerItem, .imd-picker-date-value') || alvo;
                            btn.click();
                            setTimeout(() => btn.click(), 300);
                            return true;
                        }}
                        return false;
                    }}''')
                    
                    if clicou:
                        print(f"[{sigla}] -> Duplo clique executado no dia {dia_alvo} (Tickets).")
                        page.wait_for_timeout(1500)
                    
                    btn_confirm = dialog.locator('.imd-picker-btn-confirm, button:has-text("Confirm"), .imd-picker-btn-confirm').first
                    if btn_confirm.is_visible():
                        btn_confirm.click(force=True)
                        page.wait_for_timeout(1000)
                    else:
                        page.keyboard.press("Escape")
                    break
            except: continue

        # 3. CLICA NO BOTÃO SEARCH IMEDIATAMENTE APÓS SELECIONAR A DATA
        page.wait_for_timeout(1000)
        pesquisou = False
        for frame in [page] + page.frames:
            try:
                btn_search = frame.locator('.search-btn, button:has-text("Search"), button:has-text("Pesquisar"), .ant-btn-primary').first
                if btn_search.is_visible(timeout=2000):
                    btn_search.click(force=True)
                    pesquisou = True
                    print(f"[{sigla}] -> Botão SEARCH acionado!")
                    break
            except: continue
            
        if not pesquisou:
            page.evaluate('''() => {
                let btns = Array.from(document.querySelectorAll('button'));
                let searchBtn = btns.find(b => b.innerText.includes('Search') || b.innerText.includes('Pesquisar'));
                if(searchBtn) searchBtn.click();
            }''')

        page.wait_for_timeout(8000)

        # 4. CONFIGURA PAGINAÇÃO PARA 100
        for frame in page.frames:
            try:
                caixa_paginacao = frame.locator('.imile-pagination-options .imile-select-selector').first
                if caixa_paginacao.is_visible(timeout=1000):
                    if "100" not in frame.locator('.imile-pagination-options .imile-select-selection-item').first.inner_text():
                        caixa_paginacao.click(force=True)
                        page.wait_for_timeout(1000)
                        opcao_100 = frame.locator('div[title="100 / página"], div[title="100 / page"]').first
                        if opcao_100.is_visible(timeout=2000):
                            opcao_100.click(force=True)
                            page.wait_for_timeout(8000) 
                    break 
            except: continue

        # 5. CONTA AS LINHAS DA TABELA
        total_alvo = 0
        continua_buscando = True
        paginas_lidas = 0

        while continua_buscando and paginas_lidas < 30:
            paginas_lidas += 1
            for frame in page.frames:
                try:
                    resultado = frame.evaluate('''() => {
                        const headers = Array.from(document.querySelectorAll("thead th")).map(h => h.innerText.trim().toLowerCase());
                        const idx = headers.findIndex(h => h.includes("create") || h.includes("criaç") || h.includes("criac"));
                        if (idx === -1) return { datas: [], erro: "Sem coluna" };
                        const rows = Array.from(document.querySelectorAll("tr.imile-table-row"));
                        const datas = rows.map(row => {
                            const cells = row.querySelectorAll("td");
                            return cells[idx] ? cells[idx].innerText.trim() : "";
                        }).filter(d => d !== "");
                        return { datas: datas, erro: null };
                    }''')

                    if resultado and not resultado.get("erro"):
                        datas_pagina = resultado["datas"]
                        for data_texto in datas_pagina:
                            if data_texto.startswith(data_alvo):
                                total_alvo += 1
                            elif data_texto < data_alvo:
                                continua_buscando = False
                                break

                        if len(datas_pagina) < 100: continua_buscando = False

                        if continua_buscando:
                            btn_next = frame.locator('.imile-pagination-next, li[title="Next Page"], li[title="Próxima Página"]').first
                            if btn_next.is_visible():
                                if "disabled" in (btn_next.get_attribute("class") or "") or btn_next.get_attribute("aria-disabled") == "true":
                                    continua_buscando = False
                                else:
                                    btn_next.click(force=True)
                                    page.wait_for_timeout(5000)
                            else: continua_buscando = False
                        break
                except: continue
        return total_alvo
    except Exception as e:
        print(f"[{sigla}] ❌ Erro ao extrair Acareações: {e}")
        return 0

def rodar_retroativo(sigla, usuario, senha, lista_datas, p):
    print(f"\n{'='*50}\n⏳ INICIANDO BACKFILL: {sigla}\n{'='*50}")
    if not usuario or not senha: return

    browser = p.chromium.launch(headless=False, args=["--start-maximized"])
    context = browser.new_context(no_viewport=True)
    page = context.new_page()

    try:
        page.goto("https://ds-login.imile.com/", wait_until="domcontentloaded")
        page.fill('input[type="text"]', usuario)
        page.fill('input[type="password"]', senha)
        page.get_by_role("button", name=re.compile(r"登录|Conecte-se|Login", re.IGNORECASE)).click()
        page.wait_for_timeout(5000)
        try: page.locator('.close-icon').first.click(force=True)
        except: pass

        for data_alvo in lista_datas:
            entregues = extrair_entregas_imile(page, sigla, data_alvo)
            acareacoes = extrair_acareacoes_imile(page, sigla, data_alvo)
            salvar_no_historico(sigla, entregues, acareacoes, data_alvo)

    except Exception as e: print(f"❌ Erro crítico: {e}")
    finally: browser.close()

def main():
    bases = [
        ("JML", os.getenv("IMILE_JML_USER"), os.getenv("IMILE_JML_PASS")),
        ("ITR", os.getenv("IMILE_ITR_USER"), os.getenv("IMILE_ITR_PASS")),
        ("CTG", os.getenv("IMILE_CTG_USER"), os.getenv("IMILE_CTG_PASS"))
    ]

    hoje = datetime.now()
    primeiro_dia = hoje.replace(day=1)
    ontem = hoje - timedelta(days=1)
    
    lista_datas = []
    atual = primeiro_dia
    while atual <= ontem:
        lista_datas.append(atual.strftime("%Y-%m-%d"))
        atual += timedelta(days=1)

    print(f"🚀 Iniciando extração retroativa do dia {lista_datas[0]} ao dia {lista_datas[-1]}...")

    with sync_playwright() as p:
        for sigla, usuario, senha in bases:
            rodar_retroativo(sigla, usuario, senha, lista_datas, p)
            
    print("\n🏁 DADOS RETROATIVOS CONCLUÍDOS! Seu Dashboard no site já está atualizado.")

if __name__ == "__main__":
    main()