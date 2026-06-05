print("   -> A carregar módulos básicos...")
import os
import unicodedata
import re
from datetime import datetime, timedelta

print("   -> A carregar o Pandas (Isto pode demorar na 1ª vez)...")
import pandas as pd

print("   -> A carregar o Playwright...")
from playwright.sync_api import sync_playwright

print("   -> A carregar dependências do SLAimile...")
from SLAimile import baixar_inventario
from imile_utils import login_imile
from config import ITR_LOC_CITIES, ITR_INT_CITIES, JML_LOC_CITIES, JML_INT_CITIES, BACKLOG_FIM_DE_SEMANA, BACKLOG_DIA_DE_SEMANA

print("   -> Módulos do gerar_perdas carregados com sucesso!")

def remover_acentos(texto):
    if pd.isna(texto): return ""
    texto = str(texto).upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

def extrair_relatorio_diario_completo(usuario_jml, senha_jml, usuario_itr, senha_itr):
    hoje = datetime.now()
    dia_semana_atual = hoje.weekday()
    
    if dia_semana_atual == 0:
        data_alvo = hoje - timedelta(days=2)
        nome_dia = "SÁBADO"
        backlog_config = BACKLOG_FIM_DE_SEMANA
    else:
        data_alvo = hoje - timedelta(days=1)
        dias = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO", "DOMINGO"]
        nome_dia = dias[data_alvo.weekday()]
        backlog_config = BACKLOG_DIA_DE_SEMANA
    
    backlog_int = backlog_config["int"]
    backlog_loc = backlog_config["loc"]
        
    data_str_br = data_alvo.strftime("%d/%m")
    data_alvo_imile = data_alvo.strftime("%Y-%m-%d")

    TRADUCOES_STATUS = {
        "assign da": "ATRIBUÍDO",
        "assigned": "ATRIBUÍDO",
        "atribuido": "ATRIBUÍDO",
        "atribuído": "ATRIBUÍDO",
        "out for delivery": "EM ROTA",
        "em rota de entrega": "EM ROTA",
        "delivering": "EM ROTA",
        "driver inventory": "Driver Inventory",
        "arrive": "Recebido no DS",
        "arrived at delivery station": "Recebido no DS",
        "arrived at ds": "Recebido no DS",
        "receive": "Recebido",
        "received": "Recebido",
        "offloading": "Descarregado",
        "unloaded": "Descarregado",
        "back to warehouse": "Return",
        "return": "Return",
        "returned": "Return"
    }

    def processar_base_completa(sigla, usuario, senha, cidades_loc, cidades_int, p):
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(accept_downloads=True, no_viewport=True)
        page = context.new_page()
        
        try:
            print(f"\n[{sigla}] Iniciando extração completa (Inventário + Acareações)...")
            login_imile(page, usuario, senha)
            caminho_arquivo = baixar_inventario(page, sigla)
            
            df_int_counts, df_loc_counts, total_perdas = {}, {}, 0
            
            if caminho_arquivo and os.path.exists(caminho_arquivo):
                df = pd.read_excel(caminho_arquivo, sheet_name=0)
                col_city = next((c for c in df.columns if 'destination city' in c.lower()), 'Destination City')
                col_backlog = next((c for c in df.columns if 'backlog' in c.lower()), 'Backlog time(Station)')
                col_status = next((c for c in df.columns if 'last scan type' in c.lower()), 'Last Scan Type')
                
                df[col_city] = df[col_city].apply(remover_acentos)
                df[col_backlog] = pd.to_numeric(df[col_backlog], errors='coerce').fillna(0)
                df[col_status] = df[col_status].fillna("").astype(str)
                
                df_int = df[(df[col_city].isin(cidades_int)) & (df[col_backlog] == backlog_int)].copy()
                df_loc = df[(df[col_city].isin(cidades_loc)) & (df[col_backlog] == backlog_loc)].copy()
                
                def formatar_status(st):
                    st_lower = str(st).lower().strip()
                    return TRADUCOES_STATUS.get(st_lower, str(st).title())

                df_int['Status_Formatado'] = df_int[col_status].apply(formatar_status)
                df_loc['Status_Formatado'] = df_loc[col_status].apply(formatar_status)
                
                df_int_counts = df_int['Status_Formatado'].value_counts().to_dict()
                df_loc_counts = df_loc['Status_Formatado'].value_counts().to_dict()
                total_perdas = len(df_int) + len(df_loc)

                if sigla == "JML":
                    df_loc.to_excel("DEBUG_JML_LOC.xlsx", index=False)

            page.get_by_text("Service", exact=False).first.click(force=True)
            page.wait_for_timeout(2000)
            page.get_by_text("Service", exact=False).first.click(force=True)
            page.wait_for_timeout(1000)
            page.get_by_text("Ticket management", exact=False).first.click(force=True)
            page.wait_for_timeout(1000)
            page.get_by_text("Complaint ticket", exact=False).first.click(force=True)
            page.wait_for_timeout(10000) 

            for frame in page.frames:
                try:
                    caixa_paginacao = frame.locator('.imile-pagination-options .imile-select-selector').first
                    if caixa_paginacao.is_visible(timeout=1000):
                        if "100" not in frame.locator('.imile-pagination-options .imile-select-selection-item').first.inner_text():
                            caixa_paginacao.click(force=True)
                            page.wait_for_timeout(1500)
                            opcao_100 = frame.locator('div[title="100 / página"], div[title="100 / page"]').first
                            if opcao_100.is_visible():
                                opcao_100.click(force=True)
                                page.wait_for_timeout(10000) 
                        break 
                except: continue
            
            page.wait_for_timeout(5000)

            qtd_acareacoes = 0
            for frame in page.frames:
                try:
                    if "Ticket" in frame.content() or "Complaint" in frame.content():
                        res = frame.evaluate(f'''() => {{
                            let count = 0;
                            const rows = document.querySelectorAll("tr.imile-table-row");
                            const headers = Array.from(document.querySelectorAll("thead th")).map(h => h.innerText.trim());
                            const createTimeIndex = headers.findIndex(h => h.includes("Create Time") || h.includes("Cria"));
                            
                            if (createTimeIndex >= 0) {{
                                rows.forEach(row => {{
                                    const cells = Array.from(row.querySelectorAll("td"));
                                    if (cells[createTimeIndex]) {{
                                        if (cells[createTimeIndex].innerText.trim().includes("{data_alvo_imile}")) {{
                                            count++;
                                        }}
                                    }}
                                }});
                            }}
                            return count;
                        }}''')
                        if res: qtd_acareacoes += res
                except: continue

            return df_int_counts, df_loc_counts, total_perdas, qtd_acareacoes

        except Exception as e:
            print(f"[{sigla}] Erro no fluxo: {e}")
            return {}, {}, 0, "Erro"
        finally:
            browser.close()

    with sync_playwright() as p:
        itr_int_counts, itr_loc_counts, itr_tot, itr_aca = processar_base_completa("ITR", usuario_itr, senha_itr, ITR_LOC_CITIES, ITR_INT_CITIES, p)
        jml_int_counts, jml_loc_counts, jml_tot, jml_aca = processar_base_completa("JML", usuario_jml, senha_jml, JML_LOC_CITIES, JML_INT_CITIES, p)

        # --- GERAÇÃO DA IMAGEM ÚNICA (Painel Consolidado) ---
        def montar_tabela(titulo, counts):
            html = f"<div><table><tr><th class='left'>{titulo}</th><th class='right'>QTD</th></tr>"
            for st, qtd in counts.items():
                html += f"<tr><td class='left'>{st}</td><td>{qtd}</td></tr>"
            if not counts:
                html += "<tr><td class='left' colspan='2'>Nenhum pacote</td></tr>"
            html += f"<tr class='total'><td class='left'>Total Geral</td><td>{sum(counts.values())}</td></tr></table></div>"
            return html

        html_unico = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {{ background: white; margin: 0; padding: 20px; font-family: Calibri, Arial, sans-serif; }}
                #print-perdas {{ display: inline-flex; flex-direction: column; gap: 20px; padding: 10px; background: white; }}
                .linha {{ display: flex; gap: 20px; }}
                table {{ border-collapse: collapse; font-size: 14px; width: 300px; }}
                th.left {{ background: #203764; color: white; text-align: left; padding: 6px; border: 1px solid white; font-size: 16px; }}
                th.right {{ background: #203764; color: white; text-align: center; padding: 6px; border: 1px solid white; font-size: 16px; }}
                td {{ background: #B4C6E7; color: black; padding: 6px; border: 1px solid white; text-align: center; }}
                td.left {{ text-align: left; }}
                tr.total td {{ background: #203764; color: white; font-weight: bold; }}
            </style>
        </head>
        <body>
            <div id="print-perdas">
                <div class="linha">
                    {montar_tabela("PERDAS ITR INT", itr_int_counts)}
                    {montar_tabela("PERDAS ITR LOC", itr_loc_counts)}
                </div>
                <div class="linha">
                    {montar_tabela("PERDAS JML INT", jml_int_counts)}
                    {montar_tabela("PERDAS JML LOC", jml_loc_counts)}
                </div>
            </div>
        </body>
        </html>
        """
        
        path_img_unica = "Print_Perdas_Consolidado.png"
        try:
            browser = p.chromium.launch(headless=True)
            page_print = browser.new_page()
            page_print.set_content(html_unico)
            page_print.locator("#print-perdas").wait_for(state="visible", timeout=10000)
            page_print.locator("#print-perdas").screenshot(path=path_img_unica)
        except Exception as e:
            print("Erro ao gerar imagem única:", e)
            path_img_unica = None
        finally:
            browser.close()

    # TEXTO PARA O WHATSAPP
    msg = f"BOM DIA\nSEGUE PERDAS E ACAREAÇÕES ({nome_dia} {data_str_br})\n\n"
    
    msg += f"🏢 BASE ITR (Perdas: {itr_tot} | Acareações de Ontem: {itr_aca})\n\n"
    msg += f"🔹 ITR INT \n"
    for st, qtd in itr_int_counts.items(): msg += f"* {st}: {qtd}\n"
    if not itr_int_counts: msg += "* Nenhum\n"
    msg += f"\n🔹 ITR LOC \n"
    for st, qtd in itr_loc_counts.items(): msg += f"* {st}: {qtd}\n"
    if not itr_loc_counts: msg += "* Nenhum\n"

    msg += f"\n\n🏢 BASE JML (Perdas: {jml_tot} | Acareações de Ontem: {jml_aca})\n\n"
    msg += f"🔹 JML INT \n"
    for st, qtd in jml_int_counts.items(): msg += f"* {st}: {qtd}\n"
    if not jml_int_counts: msg += "* Nenhum\n"
    msg += f"\n🔹 JML LOC \n"
    for st, qtd in jml_loc_counts.items(): msg += f"* {st}: {qtd}\n"
    if not jml_loc_counts: msg += "* Nenhum\n"

    # Retornamos a imagem consolidada no primeiro slot, e 'None' no segundo para o bot enviar só 1 foto
    return msg, path_img_unica, None