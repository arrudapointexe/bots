import os
import re
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from imile_utils import login_imile
from google_utils import get_gspread_client
from config import BACKLOG_CRITICO

load_dotenv()

# Dicionário de tradução dos Status (Inglês -> Português)
TRADUCOES_STATUS = {
    "delivering": "Em rota de entrega",
    "out for delivery": "Em rota de entrega",
    "assign da": "atribuído",
    "assigned": "atribuído",
    "offloading": "descarregado",
    "unloaded": "descarregado",
    "arrive": "recebido no DS",
    "arrived at delivery station": "recebido no DS",
    "arrived at ds": "recebido no DS",
    "receive": "Recebido",
    "received": "Recebido",
    "back to warehouse": "Return",
    "return": "Return",
    "returned": "Return"
}

# ==============================================================
# FUNÇÃO: ATUALIZAR PLANILHA DE URGÊNCIAS NO SHEETS
# ==============================================================
def atualizar_planilha_urgencias(df_nova, sigla):
    print("\n" + "="*50)
    print(f"☁️ ATUALIZANDO FILA DE URGÊNCIAS NO SHEETS ({sigla})...")
    try:
        cliente = get_gspread_client()
        planilha = cliente.open(os.getenv("NOME_PLANILHA", "acareaBase"))

        nome_aba = "Urgencias_Latencia"
        try:
            aba = planilha.worksheet(nome_aba)
            dados_existentes = aba.get_all_values()
        except gspread.exceptions.WorksheetNotFound:
            aba = planilha.add_worksheet(title=nome_aba, rows="1000", cols="5")
            dados_existentes = []

        df_final = pd.DataFrame(columns=["Código de Rastreio", "Motorista", "Status", "Dias Parado", "Base"])
        
        if len(dados_existentes) > 1:
            df_existente = pd.DataFrame(dados_existentes[1:], columns=dados_existentes[0])
            if not df_existente.empty and 'Base' in df_existente.columns:
                df_final = df_existente[df_existente['Base'] != sigla]
        
        if not df_nova.empty:
            df_final = pd.concat([df_final, df_nova], ignore_index=True)

        aba.clear()
        
        if not df_final.empty:
            df_final['Dias Parado'] = pd.to_numeric(df_final['Dias Parado'], errors='coerce').fillna(0).astype(int)
            df_final = df_final.sort_values(by="Dias Parado", ascending=False)
            df_final = df_final.astype(str)
            
            lista_dados = [df_final.columns.values.tolist()] + df_final.values.tolist()
            
            try:
                aba.update(lista_dados)
            except:
                aba.update('A1', lista_dados)
                
            print(f"✅ Sucesso! Inseridos {len(df_nova)} pacotes críticos de {sigla} na aba {nome_aba}.")
        else:
            try:
                aba.update([["Código de Rastreio", "Motorista", "Status", "Dias Parado", "Base"]])
            except:
                aba.update('A1', [["Código de Rastreio", "Motorista", "Status", "Dias Parado", "Base"]])
            print(f"✅ Sucesso! Nenhum pacote crítico de {sigla} no momento. Tabela atualizada.")
            
    except Exception as e:
        print(f"❌ Erro ao atualizar o Google Sheets: {e}")

def baixar_inventario(page, sigla):
    print(f"[{sigla}] Navegando para Monitoramento -> Monitor do inventário...")
    page.wait_for_timeout(15000)
    page.get_by_text("Monitor").first.click(force=True)
    page.wait_for_timeout(3000)
    page.get_by_text("Operation Monitor").first.click(force=True)
    page.wait_for_timeout(3000)
    page.get_by_text("Inventory Monitor").first.click(force=True)
    page.wait_for_timeout(8000)

    print(f"[{sigla}] Solicitando extração do inventário...")
    page.locator('.ImileActionButton-root:has-text("Export")').first.click(force=True)
    page.wait_for_timeout(1000)
    page.locator('li.ImileMenuItem-root:has-text("Export All")').first.click(force=True)
    page.wait_for_timeout(1000)
    page.locator('.export-button').first.click(force=True)

    print(f"[{sigla}] Aguardando a iMile gerar o arquivo...")
    page.wait_for_timeout(10000)

    print(f"[{sigla}] Baixando o arquivo bruto do inventário...")
    icone_download = page.locator('span.Imile-ButtonIcon-root svg path[d*="M8 0C3.57"]').locator('..').first
    icone_download.click(force=True)

    btn_baixar = page.locator('button:has-text("download")').last
    btn_baixar.wait_for(state="visible", timeout=15000)

    with page.expect_download(timeout=120000) as download_info:
        btn_baixar.click(force=True)

    download = download_info.value
    arquivo = f"Inventario_Vencimento_{sigla}.xlsx"
    download.save_as(arquivo)

    print(f"[{sigla}] ✅ Planilha baixada com sucesso: {arquivo}")
    return arquivo

def gerar_prints_e_mensagens(caminho_arquivo, sigla, page):
    print("\n" + "="*50)
    print(f"📸 GERANDO PRINTS E TEXTOS PARA O TELEGRAM E PLANILHA - {sigla}...")
    print("="*50)

    try:
        df = pd.read_excel(caminho_arquivo)
        
        col_awb = next((col for col in df.columns if 'tracking' in str(col).lower() or 'awb' in str(col).lower() or 'waybill' in str(col).lower()), 'Tracking Number')
        col_status = next((col for col in df.columns if str(col).lower() == 'last scan type'), 'Last Scan Type')
        col_backlog = next((col for col in df.columns if str(col).lower() == 'backlog time(station)'), 'Backlog time(Station)')
        col_motorista = next((col for col in df.columns if str(col).lower() == 'delivery associate' or str(col).lower() == 'driver name'), 'Delivery Associate')

        def traduzir(valor):
            if pd.isna(valor) or str(valor).strip() == "":
                return '(vazio)'
            valor_str = str(valor).strip()
            return TRADUCOES_STATUS.get(valor_str.lower(), valor_str)

        df[col_status] = df[col_status].apply(traduzir)
        df[col_backlog] = df[col_backlog].fillna('(vazio)')

        # ==========================================
        # 0. ISOLAR A FILA DE URGÊNCIAS (>= 3 DIAS)
        # ==========================================
        df_urgencias = pd.DataFrame()
        df_lat_critica = df.copy()
        df_lat_critica[col_backlog] = pd.to_numeric(df_lat_critica[col_backlog], errors='coerce')
        df_lat_critica = df_lat_critica[df_lat_critica[col_backlog] >= BACKLOG_CRITICO].copy()
        
        if not df_lat_critica.empty:
            df_urgencias = df_lat_critica[[col_awb, col_motorista, col_status, col_backlog]].copy()
            df_urgencias['Base'] = sigla
            df_urgencias.columns = ["Código de Rastreio", "Motorista", "Status", "Dias Parado", "Base"]
            df_urgencias["Motorista"] = df_urgencias["Motorista"].apply(lambda x: re.sub(r'\(.*?\)', '', str(x)).strip())

        # ==========================================
        # 1. GERAR PRINT E TEXTO DA LATÊNCIA
        # ==========================================
        def filtro_ate_20_dias(val):
            if str(val) == '(vazio)': 
                return True
            try:
                return float(val) <= 20
            except:
                return True

        df_latencia = df[df[col_backlog].apply(filtro_ate_20_dias)].copy()

        tabela_lat = pd.crosstab(df_latencia[col_status], df_latencia[col_backlog], margins=True, margins_name="Total Geral")
        
        if "Total Geral" in tabela_lat.columns:
            total_row = tabela_lat.loc[["Total Geral"]]
            corpo = tabela_lat.drop("Total Geral").sort_values(by="Total Geral", ascending=False)
            tabela_lat = pd.concat([corpo, total_row])

        html_lat = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {{ background: white; margin: 0; padding: 20px; }}
                #print-latencia {{ display: inline-block; padding: 5px; background: white; }}
                table {{ border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 14px; }}
                .title {{ background: #203764; color: white; font-weight: bold; font-size: 16px; padding: 6px; text-align: left; }}
                th {{ background: #203764; color: white; padding: 5px 10px; border: 1px solid white; text-align: center; }}
                td {{ background: #B4C6E7; color: black; padding: 5px 10px; border: 1px solid white; text-align: center; }}
                td.left {{ text-align: left; }}
                tr.total td {{ background: #203764; color: white; font-weight: bold; }}
            </style>
        </head>
        <body>
            <div id="print-latencia">
                <table>
                <tr><td colspan="{len(tabela_lat.columns)+1}" class="title">LATENCIA IMILE {sigla}</td></tr>
                <tr><th style="text-align: left;">Rótulos de Linha</th>
        """
        for col in tabela_lat.columns: 
            html_lat += f'<th>{col}</th>'
        html_lat += '</tr>'

        msg_latencia = f"🚨 *SITUAÇÃO DE LATÊNCIA | {sigla}* 🚨\n━━━━━━━━━━━━━━━━━━━━━━\n"
        
        for idx, row in tabela_lat.iterrows():
            is_total = ' class="total"' if str(idx) == 'Total Geral' else ''
            html_lat += f'<tr{is_total}><td class="left">{idx}</td>'
            
            if str(idx) != 'Total Geral' and row['Total Geral'] > 0:
                msg_latencia += f"📌 *{idx}:* {int(row['Total Geral'])} pacotes\n"
            elif str(idx) == 'Total Geral':
                msg_latencia += f"━━━━━━━━━━━━━━━━━━━━━━\n📦 *TOTAL GERAL:* {int(row['Total Geral'])} pacotes\n"

            for col in tabela_lat.columns:
                val = row[col]
                texto = "" if (val == 0 and not is_total) else str(int(val))
                html_lat += f'<td>{texto}</td>'
            html_lat += '</tr>'
        html_lat += "</table></div></body></html>"

        # ==========================================
        # 2. GERAR PRINTS E TEXTOS DE EM ROTA
        # ==========================================
        df_rota = df[df[col_status].astype(str).str.contains('Em rota de entrega', case=False, na=False)].copy()
        
        # --- MÁGICA DA COLUNA Q (ÍNDICE 16) PARA OCORRÊNCIAS ---
        if len(df.columns) > 16:
            col_q = df.columns[16] # Coluna Q = Índice 16
            
            # Pega o valor, converte para string, tira espaços e remove o '.0' se o Excel ler como float
            val_str = df_rota[col_q].astype(str).str.strip().str.replace('.0', '', regex=False)
            
            # Regra: Tem ocorrência se o valor não for '0' e não estiver vazio/nulo
            tem_ocorrencia = (val_str != '0') & (val_str != '') & (~val_str.str.lower().isin(['nan', 'nat', 'none']))
            
            df_rota['Tipo_Rota'] = 'Limpo'
            df_rota.loc[tem_ocorrencia, 'Tipo_Rota'] = 'Ocorrencia'
        else:
            df_rota['Tipo_Rota'] = 'Limpo'
            
        tags = ['INT', 'LOC'] if sigla.upper() == 'JML' else ['STB', 'ITR']

        paths_img_rota = []
        msgs_rota = []
        page_print = page.context.new_page()

        try:
            path_img_lat = f"Print_Latencia_{sigla}.png"
            page_print.set_content(html_lat)
            page_print.locator("#print-latencia").wait_for(state="visible", timeout=10000)
            page_print.locator("#print-latencia").screenshot(path=path_img_lat)

            for tag in tags:
                df_tag = df_rota[df_rota[col_motorista].astype(str).str.contains(tag, na=False, case=False)]
                
                msg_tag = f"🚚 *STATUS DE ROTA | {sigla} - {tag}* 🚚\n━━━━━━━━━━━━━━━━━━━━━━\n"
                
                html_tag = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <style>
                        body {{ background: white; margin: 0; padding: 20px; }}
                        #print-rota {{ display: inline-block; padding: 5px; background: white; }}
                        table {{ border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 14px; margin-bottom: 20px; width: 600px; }}
                        th.left {{ background: #203764; color: #FF0000; text-align: left; padding: 5px; border: 1px solid white; }}
                        th.right {{ background: #203764; color: #FFFF00; text-align: center; padding: 5px; border: 1px solid white; }}
                        td {{ background: #B4C6E7; color: black; padding: 5px; border: 1px solid white; text-align: center; }}
                        td.left {{ text-align: left; }}
                        tr.total td {{ background: #203764; color: white; font-weight: bold; }}
                    </style>
                </head>
                <body>
                    <div id="print-rota">
                """

                if not df_tag.empty:
                    # Tabela Dinâmica que separa os Limpos das Ocorrências
                    pt = pd.pivot_table(
                        df_tag, values=col_awb, 
                        index=col_motorista, 
                        columns='Tipo_Rota', aggfunc='count', fill_value=0
                    )
                    
                    for c in ['Limpo', 'Ocorrencia']:
                        if c not in pt.columns:
                            pt[c] = 0
                            
                    pt['Total'] = pt['Limpo'] + pt['Ocorrencia']
                    pt = pt.sort_values(by='Total', ascending=False)
                    
                    html_tag += f'<table><tr><th class="left">VENCIMENTO IMILE {tag}</th><th class="right">EM ROTA</th><th class="right">OCORRÊNCIA</th><th class="right">TOTAL</th></tr>'

                    total_limpo = 0
                    total_ocor = 0
                    total_geral = 0

                    for mot_completo, row in pt.iterrows():
                        qtd_limpo = int(row['Limpo'])
                        qtd_ocor = int(row['Ocorrencia'])
                        qtd_total = int(row['Total'])
                        
                        total_limpo += qtd_limpo
                        total_ocor += qtd_ocor
                        total_geral += qtd_total

                        nome_limpo = re.sub(r'\(.*?\)', '', str(mot_completo)).strip()
                        nome_limpo = nome_limpo.split('-')[0].strip()
                        primeiro_nome = nome_limpo.split(' ')[0].strip().upper()

                        html_tag += f'<tr><td class="left">{mot_completo}</td><td>{qtd_limpo if qtd_limpo>0 else ""}</td><td>{qtd_ocor if qtd_ocor>0 else ""}</td><td>{qtd_total}</td></tr>'
                        
                        texto_telegram = f"{qtd_total} pacotes"
                        if qtd_ocor > 0:
                            texto_telegram += f" ({qtd_ocor} ocor.)"
                            
                        msg_tag += f"👤 @{primeiro_nome} ➜ {texto_telegram}\n"
                    
                    html_tag += f'<tr class="total"><td class="left">Total Geral</td><td>{total_limpo}</td><td>{total_ocor}</td><td>{total_geral}</td></tr></table>'
                    msg_tag += f"━━━━━━━━━━━━━━━━━━━━━━\n📈 *Total em Rota:* {total_geral} pacotes\n"
                else:
                    html_tag += f"<table><tr><th class='left'>VENCIMENTO IMILE {tag}</th></tr><tr><td class='left'>Nenhum pacote em rota.</td></tr></table>"
                    msg_tag += f"✅ *Tudo limpo!* Nenhum pacote pendente em rota para a base {tag}."

                html_tag += "</div></body></html>"

                path_img_tag = f"Print_Rota_{sigla}_{tag}.png"
                page_print.set_content(html_tag)
                page_print.locator("#print-rota").wait_for(state="visible", timeout=10000)
                page_print.locator("#print-rota").screenshot(path=path_img_tag)

                paths_img_rota.append(path_img_tag)
                msgs_rota.append(msg_tag)

        finally:
            page_print.close()

        print(f"✅ Imagens e textos criados com sucesso para {sigla}!")
        
        atualizar_planilha_urgencias(df_urgencias, sigla)
        
        return {
            "img_lat": path_img_lat,
            "txt_lat": msg_latencia,
            "imgs_rota": paths_img_rota,
            "txts_rota": msgs_rota,
            "df_urgencias": df_urgencias
        }

    except Exception as e:
        print(f"❌ Erro detalhado ao gerar prints para {sigla}: {e}")
        return None
        
def processar_conta(sigla, usuario, senha, p):
    print(f"\n{'='*50}\n🚀 INICIANDO CONTA: {sigla}\n{'='*50}")
    if not usuario or not senha:
        print(f"⚠️ Credenciais de {sigla} ausentes. Pulando...")
        return None

    browser = p.chromium.launch(headless=True, args=["--start-maximized"])
    context = browser.new_context(no_viewport=True, accept_downloads=True)
    page = context.new_page()

    df_critico = pd.DataFrame()

    try:
        print(f"[{sigla}] Fazendo login...")
        login_imile(page, usuario, senha)

        arquivo_bruto = baixar_inventario(page, sigla)
        if arquivo_bruto and os.path.exists(arquivo_bruto):
            dados_telegram = gerar_prints_e_mensagens(arquivo_bruto, sigla, page)
            if dados_telegram:
                df_critico = dados_telegram.get("df_urgencias", pd.DataFrame())
            
    except Exception as e:
        print(f"❌ Erro crítico na base {sigla}: {e}")

    finally:
        browser.close()
    
    return df_critico


def main():
    contas = [
        ("JML", os.getenv("IMILE_JML_USER"), os.getenv("IMILE_PASS")),
        ("ITR", os.getenv("IMILE_ITR_USER"), os.getenv("IMILE_PASS")),
        ("GNH", os.getenv("IMILE_GNH_USER"), os.getenv("IMILE_PASS")),
        ("MNT", os.getenv("IMILE_MNT_USER"), os.getenv("IMILE_PASS")),
        ("GVR", os.getenv("IMILE_GVR_USER"), os.getenv("IMILE_PASS")),
        ("TFO", os.getenv("IMILE_TFO_USER"), os.getenv("IMILE_PASS")),
        ("RBN", os.getenv("IMILE_RBN_USER"), os.getenv("IMILE_PASS")),
        ("CPH", os.getenv("IMILE_CPH_USER"), os.getenv("IMILE_PASS")),
        ("QHG", os.getenv("IMILE_QHG_USER"), os.getenv("IMILE_PASS")),
        ("CTP", os.getenv("IMILE_CTP_USER"), os.getenv("IMILE_PASS"))
    ]

    with sync_playwright() as p:
        for sigla, usuario, senha in contas:
            processar_conta(sigla, usuario, senha, p)

if __name__ == "__main__":
    main()