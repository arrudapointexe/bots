#TELEGRAM_BOT_TOKEN="8677981323:AAEbhniLvA8UaozXVcTQg1VCSwKkt95CnOw"
import os
import glob
import zipfile
import pandas as pd
from playwright.sync_api import sync_playwright

# ==========================================
# CONFIGURAÇÕES
# ==========================================
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads_shopee")
USER_DATA_DIR = "C:\\PerfilBotShopee"

MAPEAMENTO_STATUS = {
    'Hub_Received': 'Em Base',
    'Hub_Assigned': 'Expedir',
    'Hub_Assigning': 'Expedir', 
    'Delivering': 'Em Rota', 
    'OnHold': 'Ocorrencia', 
    'Delivered': 'Entregue', 
    'Completed': 'Entregue',
    'Failed Delivery': 'Ocorrencia', 
    'Cancelled': 'Ocorrencia', 
    'Return to Hub': 'Ocorrencia'
}

def limpar_pasta_downloads():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    for arquivo in glob.glob(os.path.join(DOWNLOAD_DIR, "*")):
        try: os.remove(arquivo)
        except: pass

# ==========================================
# 1. FUNÇÃO DE NAVEGAÇÃO E DOWNLOAD
# ==========================================
def baixar_planilha_shopee():
    print("🧹 Limpando downloads antigos...")
    limpar_pasta_downloads()
    
    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            accept_downloads=True
        )
        
        page = browser.pages[0]
        try:
            print("🌐 Acessando o portal da Shopee...")
            # Tiramos o networkidle e colocamos domcontentloaded (bem mais seguro e rápido)
            page.goto("https://spx.shopee.com.br/#/index", wait_until="domcontentloaded")

            print("📦 Navegando para Rastreio de Pedidos...")
            # O robô agora espera até 60 segundos O BOTÃO aparecer na tela, em vez de esperar a rede
            menu_pedidos = page.locator("span.sub-menu-title", has_text="Pedidos")
            menu_pedidos.wait_for(state="visible", timeout=60000)
            menu_pedidos.click()
            
            submenu_rastreio = page.locator("a[title='Rastreio de pedidos']")
            submenu_rastreio.wait_for(state="visible", timeout=10000)
            submenu_rastreio.click()

            print("⚙️ Abrindo Exportação Avançada...")
            # Espera a página de rastreio carregar o botão de Exportar
            btn_exportar = page.locator("button", has_text="Exportar")
            btn_exportar.wait_for(state="visible", timeout=60000)
            btn_exportar.hover()
            
            page.locator("div.ssc-dropdown-item", has_text="Exportar Pedido Avançado").click()

            print("🔎 Aplicando filtros de Status...")
            status_desejados = ["Delivering", "OnHold", "Hub_Assigned", "Hub_Assigning", "Hub_Received"]
            for status in status_desejados:
                linha_status = page.locator("div.s-tree-node__content").filter(
                    has=page.locator(f"span.ssc-tree-node__label:text-is('{status}')")
                )
                linha_status.locator("label.ssc-checkbox-wrapper").click()

            print("👤 Aplicando filtro de Conta...")
            conta_pedido_div = page.locator("div.ssc-form-item", has_text="Conta do pedido:")
            conta_pedido_div.locator("span.ssc-tree-node__label", has_text="All").first.click()

            page.locator("button.ssc-btn-type-primary", has_text="Confirmar").click()

            print("⏳ Aguardando o processamento do relatório no servidor...")
            tarefa_row = page.locator("div.task-row", has_text="Forward Order").first
            botao_baixar = tarefa_row.locator("button", has_text="Baixar")
            
            # Aqui deixei um timeout longo (5 minutos) porque a Shopee às vezes demora para gerar o Excel
            botao_baixar.wait_for(state="visible", timeout=300000) 
            
            print("⬇️ Baixando arquivo...")
            with page.expect_download() as download_info:
                botao_baixar.click()
            
            download = download_info.value
            caminho_arquivo = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
            download.save_as(caminho_arquivo)
            print(f"✅ Download concluído: {download.suggested_filename}")
            
            if caminho_arquivo.endswith('.zip'):
                print("🗜️ Extraindo arquivo ZIP...")
                with zipfile.ZipFile(caminho_arquivo, 'r') as zip_ref:
                    zip_ref.extractall(DOWNLOAD_DIR)
                arquivos = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")) + glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))
                caminho_arquivo = arquivos[0] if arquivos else caminho_arquivo

            return caminho_arquivo
        
        except Exception as e:
            print(f"❌ Erro na extração da Shopee: {e}")
            return None
        finally:
            browser.close()

# ==========================================
# 2. FUNÇÃO DE LEITURA DA BASE DE CEPS
# ==========================================
def carregar_dicionario_ceps(caminho_excel):
    print("Lendo base de CEPs (Sua planilha amarela)...")
    try:
        df_ceps = pd.read_excel(caminho_excel, dtype=str)
        df_ceps.columns = df_ceps.columns.str.strip()
        
        if 'CEPs' in df_ceps.columns and 'Responsavel' in df_ceps.columns:
            col_cep = 'CEPs'
            col_cidade = 'Responsavel' 
        else:
            col_cep = df_ceps.columns[0]
            col_cidade = df_ceps.columns[2]

        df_ceps[col_cep] = df_ceps[col_cep].str.replace(r'\D', '', regex=True)
        mapa_ceps = dict(zip(df_ceps[col_cep], df_ceps[col_cidade]))
        print(f"✅ Base de CEPs carregada com {len(mapa_ceps)} registros!")
        return mapa_ceps
    except Exception as e:
        print(f"❌ Erro ao ler a base de CEPs: {e}")
        return {}

# ==========================================
# 3. LEITURA BLINDADA (CORRIGE O CSV)
# ==========================================
def ler_planilha_shopee_segura(caminho):
    if caminho.endswith('.csv'):
        tentativas = [
            (',', 'utf-8-sig'), (',', 'utf-8'), (',', 'latin1'),
            (';', 'utf-8-sig'), (';', 'latin1')
        ]
        for sep, enc in tentativas:
            try:
                df = pd.read_csv(caminho, sep=sep, encoding=enc, dtype=str)
                if len(df.columns) > 2:
                    return df
            except:
                continue
    else:
        try: return pd.read_excel(caminho, dtype=str)
        except: pass
        
    return pd.DataFrame()

# ==========================================
# 4. TRATAMENTO DE DADOS E GERAÇÃO DA FOTO
# ==========================================
def gerar_relatorio_shopee(caminho_forward, mapa_ceps):
    print("Processando arquivo Export Forward...")
    
    df = ler_planilha_shopee_segura(caminho_forward)
    if df.empty:
        return ["❌ Erro ao ler a planilha da Shopee."], []

    # ======== O CURATIVO DA COLUNA DE DATA/HORA ESTÁ AQUI ========
    col_station = next((c for c in df.columns if c.strip().lower() in ['current station', 'estação atual', 'estacao atual']), None)
    if not col_station:
        col_station = next((c for c in df.columns if ('station' in c.lower() or 'estação' in c.lower() or 'estacao' in c.lower()) and 'time' not in c.lower() and 'hora' not in c.lower()), None)
    # ==============================================================

    if not col_station:
        colunas_achadas = ", ".join(list(df.columns)[:10])
        return [f"❌ Coluna 'Current Station' não encontrada!\n\nColunas lidas: {colunas_achadas}"], []
    
    df_filtrado = df[df[col_station].astype(str).str.contains('Monlevade|JML', na=False, case=False)]
    if df_filtrado.empty:
        bases_encontradas = df[col_station].dropna().unique()[:15]
        lista_bases = "\n- ".join(map(str, bases_encontradas))
        return [f"⚠️ Nenhum pacote de Monlevade encontrado no arquivo.\nBases achadas na coluna:\n- {lista_bases}"], []

    df = df_filtrado

    colunas_para_traduzir = {
        col_station: 'Hub',
        next((c for c in df.columns if 'zipcode' in c.lower() or 'cep' in c.lower()), 'CEP'): 'CEP', 
        next((c for c in df.columns if 'driver name' in c.lower() or 'motorista' in c.lower()), 'Responsavel'): 'Responsavel',
        next((c for c in df.columns if 'tracking number' in c.lower() or 'rastreamento' in c.lower()), 'Tracking Number'): 'Tracking Number',
        next((c for c in df.columns if 'status' in c.lower()), 'Status'): 'Status'
    }
    df.rename(columns=colunas_para_traduzir, inplace=True)

    df['Hub'] = 'XPT-JML'
    
    if 'Responsavel' in df.columns:
        df['Responsavel'] = df['Responsavel'].fillna('(vazio)')
        df.loc[df['Responsavel'].str.strip() == '', 'Responsavel'] = '(vazio)'
    
    if 'Status' in df.columns:
        df['Status'] = df['Status'].replace(MAPEAMENTO_STATUS)
    
    # Mapeamento de Cidades e Resgate de CEPs Faltantes
    ceps_fujoes = []
    if 'CEP' in df.columns:
        df['CEP'] = df['CEP'].astype(str).str.replace(r'\D', '', regex=True)
        df['Cidade'] = df['CEP'].map(mapa_ceps)
        
        ceps_fujoes = df[df['Cidade'].isna()]['CEP'].dropna().unique().tolist()
        df['Cidade'] = df['Cidade'].fillna('OUTROS_CEPS')

    rotas_shopee = {
        "Joao_Monlevade": ['João Monlevade'],
        "Barao_SB_Catas": ['Barão de Cocais', 'Santa Bárbara', 'Catas Altas'],
        "Interior": ['Nova Era', 'Bela Vista de Minas', 'São Gonçalo do Rio Abaixo', 'Rio Piracicaba', 'São Domingos do Prata', 'São José do Goiabal', 'Dionísio'],
        "Problemas_CEP": ['OUTROS_CEPS']
    }

    ordem_status = ['Em Base', 'Expedir', 'Em Rota', 'Ocorrencia', 'Entregue']
    
    imagens_geradas = []
    mensagens = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for nome_rota, cidades in rotas_shopee.items():
            df_rota = df[df['Cidade'].str.upper().isin([c.upper() for c in cidades])]
            if df_rota.empty:
                continue

            pt = pd.pivot_table(
                df_rota, values='Tracking Number', 
                index=['Hub', 'Cidade', 'Responsavel'], 
                columns=['Status'], aggfunc='count', fill_value=0
            )
            
            for col in ordem_status:
                if col not in pt.columns:
                    pt[col] = 0
                    
            pt = pt[ordem_status]
            pt['Total Geral'] = pt.sum(axis=1)

            html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {{ background: #002030; margin: 0; padding: 20px; font-family: Calibri, Arial, sans-serif; }}
                    #print-shopee {{ display: inline-block; background: #87CEFA; padding: 2px; border: 2px solid black; }}
                    table {{ border-collapse: collapse; font-size: 14px; width: 800px; }}
                    th {{ color: white; padding: 6px; text-align: center; border: 1px solid #555; background: #002030; }}
                    th.titulo {{ color: red; font-size: 20px; text-align: left; border: none; padding-bottom: 10px; }}
                    th.subtitulo {{ text-align: left; }}
                    td {{ padding: 4px 6px; border: 1px solid #777; text-align: center; color: black; font-weight: bold; }}
                    td.left {{ text-align: left; }}
                    tr.hub td {{ background: #A9D0F5; font-weight: bold; }}
                    tr.city td {{ background: #87CEFA; font-weight: bold; }}
                    tr.mot td {{ background: #87CEFA; font-weight: normal; }}
                    tr.total td {{ background: #002030; color: white; font-weight: bold; }}
                </style>
            </head>
            <body>
                <div id="print-shopee">
                    <table>
                        <tr><th colspan="7" class="titulo">SLA-Expedição-SHOPEE ({nome_rota.replace('_', ' ')})</th></tr>
                        <tr>
                            <th class="subtitulo">Responsavel</th>
                            <th>Em Base</th><th>Expedir</th><th>Em Rota</th>
                            <th>Ocorrencia</th><th>Entregue</th><th>Total Geral</th>
                        </tr>
            """

            total_absoluto = {s: 0 for s in ordem_status + ['Total Geral']}

            for hub in df_rota['Hub'].unique():
                df_hub = df_rota[df_rota['Hub'] == hub]
                totais_hub = {s: df_hub[df_hub['Status'] == s].shape[0] for s in ordem_status}
                totais_hub['Total Geral'] = sum(totais_hub.values())
                for k in total_absoluto: total_absoluto[k] += totais_hub.get(k, 0)

                html += f"<tr class='hub'><td class='left'>⊟ {hub}</td>"
                for s in ordem_status + ['Total Geral']:
                    val = totais_hub[s]
                    html += f"<td>{val if val > 0 else ''}</td>"
                html += "</tr>"

                for cidade in df_hub['Cidade'].unique():
                    df_cid = df_hub[df_hub['Cidade'] == cidade]
                    totais_cid = {s: df_cid[df_cid['Status'] == s].shape[0] for s in ordem_status}
                    totais_cid['Total Geral'] = sum(totais_cid.values())

                    html += f"<tr class='city'><td class='left' style='padding-left: 15px;'>⊟ {cidade}</td>"
                    for s in ordem_status + ['Total Geral']:
                        val = totais_cid[s]
                        html += f"<td>{val if val > 0 else ''}</td>"
                    html += "</tr>"

                    if (hub, cidade) in pt.index:
                        df_mots = pt.loc[(hub, cidade)].sort_values(by='Total Geral', ascending=False)
                        for mot, row in df_mots.iterrows():
                            html += f"<tr class='mot'><td class='left' style='padding-left: 30px;'>{mot}</td>"
                            for s in ordem_status + ['Total Geral']:
                                val = row[s]
                                html += f"<td>{int(val) if val > 0 else ''}</td>"
                            html += "</tr>"

            html += "<tr class='total'><td class='left'>Total Geral</td>"
            for s in ordem_status + ['Total Geral']:
                html += f"<td>{total_absoluto[s]}</td>"
            html += "</tr></table></div></body></html>"

            caminho_imagem = f"Print_Shopee_{nome_rota}.png"
            page.set_content(html)
            page.locator("#print-shopee").wait_for(state="visible")
            page.locator("#print-shopee").screenshot(path=caminho_imagem)

            imagens_geradas.append(caminho_imagem)
            mensagens.append(f"✅ *Shopee - Vencimentos:* {nome_rota.replace('_', ' ')}")

        browser.close()

    # Adiciona o aviso dos CEPs no final (sem formatação que buga o Telegram)
    if ceps_fujoes:
        lista = "\n".join([f"- {cep}" for cep in ceps_fujoes])
        aviso = f"⚠️ CEPs NÃO ENCONTRADOS!\nTem pacotes indo para CEPs que não existem na sua planilha base_ceps.xlsx.\nAdicione-os lá para sumir com a aba OUTROS_CEPS:\n{lista}"
        mensagens.append(aviso)

    if not imagens_geradas:
        return ["✅ O arquivo foi lido, mas não há nenhum pacote pendente nas rotas especificadas."], []

    return mensagens, imagens_geradas
