import os
import time
import pandas as pd
import re
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from imile_utils import login_imile
from google_utils import get_gspread_client

load_dotenv()

NOME_PLANILHA = os.getenv("NOME_PLANILHA", "acareaBase")

def enviar_para_google_sheets(df_final, sigla, bot=None, chat_id=None):
    msg = f"☁️ Enviando dados para a aba {sigla} no Google Sheets..."
    print(msg)
    if bot and chat_id: bot.send_message(chat_id, msg)
    
    MAX_TENTATIVAS = 3
    for tentativa in range(MAX_TENTATIVAS):
        try:
            cliente = get_gspread_client()
            planilha_mestre = cliente.open(NOME_PLANILHA)
            
            # Cria a aba automaticamente se não existir
            try:
                planilha = planilha_mestre.worksheet(sigla)
            except gspread.exceptions.WorksheetNotFound:
                msg_nova_aba = f"✨ Aba '{sigla}' não encontrada. Criando nova aba automaticamente..."
                print(msg_nova_aba)
                if bot and chat_id: bot.send_message(chat_id, msg_nova_aba)
                planilha = planilha_mestre.add_worksheet(title=sigla, rows="1000", cols="20")
            
            planilha.clear()
            dados_para_nuvem = [df_final.columns.values.tolist()] + df_final.fillna("").values.tolist()
            planilha.update('A1', dados_para_nuvem)
            
            msg_sucesso = f"✅ SUCESSO! Dados da base {sigla} atualizados na nuvem."
            print(msg_sucesso)
            if bot and chat_id: bot.send_message(chat_id, msg_sucesso)
            return  # Se deu certo, sai da função
            
        except Exception as e:
            msg_erro = f"⚠️ Falha na tentativa {tentativa + 1} de enviar {sigla} para a nuvem. Tentando novamente..."
            print(msg_erro)
            time.sleep(5) # Espera 5 segundos antes de tentar de novo
            
    # Se tentou 3 vezes e falhou em todas
    msg_falha = f"❌ Erro crítico: Não foi possível enviar a base {sigla} após {MAX_TENTATIVAS} tentativas."
    print(msg_falha)
    if bot and chat_id: bot.send_message(chat_id, msg_falha)


def rodar_automacao_acareacao(sigla, usuario, senha, bot=None, chat_id=None):
    
    def log(msg):
        texto = f"[{sigla} - Acareação] {msg}"
        print(texto)
        if bot and chat_id:
            bot.send_message(chat_id, texto)

    log("Iniciando automação...")

    with sync_playwright() as p:
        # O argumento '--disable-blink-features=AutomationControlled' é a mágica que esconde que é um robô
        browser = p.chromium.launch(
            headless=False, 
            args=["--disable-blink-features=AutomationControlled"]
        )
        
        # 1. CORREÇÃO DA TELA CORTADA: Forçamos 1920x1080 e idioma em Inglês
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            locale="en-US",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            accept_downloads=True
        )
        page = context.new_page()

        try:
            log("Acessando o portal...")
            login_imile(page, usuario, senha)
            
           # 3. NAVEGAÇÃO BLINDADA
            log("Navegando nos menus: Ticket de reclamação...")
            page.get_by_text(re.compile(r"Service|Atendimento|Serviço", re.IGNORECASE)).first.click(force=True)
            page.wait_for_timeout(1000)
            page.get_by_text(re.compile(r"Ticket management|Gestão de tickets", re.IGNORECASE)).first.click(force=True)
            page.wait_for_timeout(1000)
            page.get_by_text(re.compile(r"Complaint ticket|Ticket de reclamação|Reclamaç", re.IGNORECASE)).first.click(force=True)
            
            # 🔥 AUMENTADO PARA 25 SEGUNDOS (O site da iMile às vezes engasga para carregar a tabela)
            page.wait_for_timeout(25000) 

            # MUDAR PAGINAÇÃO PARA 100
            log("Ajustando paginação para 100...")
            for frame in page.frames:
                try:
                    caixa_paginacao = frame.locator('.imile-pagination-options .imile-select-selector').first
                    if caixa_paginacao.is_visible(timeout=3000):
                        texto_atual = frame.locator('.imile-pagination-options .imile-select-selection-item').first.inner_text()
                        if "100" not in texto_atual:
                            caixa_paginacao.click(force=True)
                            page.wait_for_timeout(2000)
                            opcao_100 = frame.locator('div[title="100 / página"], div[title="100 / page"]').first
                            if opcao_100.is_visible(timeout=4000):
                                opcao_100.click(force=True)
                                log("Paginação alterada! Aguardando site carregar as 100 linhas...")
                                # 🔥 AUMENTADO PARA 30 SEGUNDOS
                                page.wait_for_timeout(30000) 
                        break 
                except: continue
            
            # Mais uma folguinha extra antes de ler
            page.wait_for_timeout(8000)

            # BUSCA DE CÓDIGOS EM PROCESSING
            log("Buscando códigos em Processing...")
            codigos_prazos = []
            for index, frame in enumerate(page.frames):
                try:
                    if "Ticket" in frame.content():
                        res = frame.evaluate(r'''() => {
                            const itens = [];
                            const rows = document.querySelectorAll("tr.imile-table-row");
                            const headers = Array.from(document.querySelectorAll("thead th")).map(h => h.innerText.trim());

                            const awbIndex = headers.findIndex(h => h.includes("Related AWB") || h.includes("Waybill"));
                            const prazoIndex = headers.findIndex(h => h.includes("Deadline Process Time"));
                            const statusIndex = headers.findIndex(h => h.includes("Handling Status") || h.includes("Status"));
                            // 👇 NOVO: Buscando a coluna de Subtipo
                            const subTypeIndex = headers.findIndex(h => h.includes("Ticket Sub Type") || h.includes("Sub Type"));

                            rows.forEach(row => {
                                const cells = Array.from(row.querySelectorAll("td"));
                                if (!cells || cells.length === 0) return;

                                let status = "";
                                if (statusIndex >= 0 && cells[statusIndex]) {
                                    status = cells[statusIndex].innerText.trim();
                                } else {
                                    status = cells.map(c => c.innerText.trim()).includes("Processing") ? "Processing" : "";
                                }

                                if (status === "Processing") {
                                    let codigo = "";
                                    let prazo = "N/A";
                                    let sub_tipo = "N/A";

                                    if (awbIndex >= 0 && cells[awbIndex]) {
                                        codigo = cells[awbIndex].innerText.trim();
                                    } else {
                                        const match = cells.map(c => c.innerText.trim()).join(' ').match(/\b\d{13,15}\b/);
                                        if (match) codigo = match[0];
                                    }

                                    if (prazoIndex >= 0 && cells[prazoIndex]) {
                                        prazo = cells[prazoIndex].innerText.trim();
                                    }
                                    
                                    // 👇 NOVO: Extraindo o texto do Subtipo
                                    if (subTypeIndex >= 0 && cells[subTypeIndex]) {
                                        sub_tipo = cells[subTypeIndex].innerText.trim();
                                    }

                                    // 👇 ATUALIZADO: Guardando o subtipo junto com o AWB
                                    if (codigo) itens.push({ codigo, prazo, sub_tipo });
                                }
                            });
                            return itens;
                        }''')
                        if res: codigos_prazos.extend(res)
                except: continue

            codigos_map = {}
            for item in codigos_prazos:
                codigo = str(item.get('codigo', '')).strip()
                if codigo: codigos_map[codigo] = item

            codigos_finais = list(codigos_map.values())
            log(f"Total de códigos Processing capturados: {len(codigos_finais)}")

            if not codigos_finais:
                log("Nenhum código para detalhar. Finalizando...")
                browser.close()
                return None

            log("Iniciando consultas detalhadas (isso pode demorar um pouco)...")
            page.goto("https://ds.imile.com/#/DSOperation/WaybillManagement/dsTrackQuery", wait_until="domcontentloaded")
            
            dados_finais_excel = []

            for item in codigos_finais:
                codigo = item['codigo']
                prazo_da_linha = item['prazo']
                sub_tipo_da_linha = item.get('sub_tipo', 'N/A') # 👇 NOVO: Puxando o subtipo
                target_f = None
                
                for _ in range(10):
                    for f in page.frames:
                        try:
                            el = f.locator('.search-input input, input[placeholder*="insira"]').first
                            if el.is_visible(timeout=500):
                                el.fill(str(codigo))
                                f.locator('.search-btn, button:has-text("Pesquisar")').first.click(force=True)
                                target_f = f
                                break
                        except: continue
                    if target_f: break
                    page.wait_for_timeout(1000)

                if target_f:
                    page.wait_for_timeout(3000)
                    
                    try: target_f.get_by_role("tab", name="CUSTOMER INFO").first.click(force=True)
                    except: pass
                    page.wait_for_timeout(1000) 
                    
                    try: 
                        olho = target_f.locator('.detail-item', has_text="Customer Name").locator('svg, [role="button"]').first
                        if olho.is_visible(timeout=1000):
                            olho.click(force=True)
                            page.wait_for_timeout(1500)
                    except: pass

                    try: 
                        olho_tel = target_f.locator('.detail-item', has_text="Customer phone").locator('svg, [role="button"]').first
                        if olho_tel.is_visible(timeout=1000):
                            olho_tel.click(force=True)
                            page.wait_for_timeout(1500)
                    except: pass
                    
                    dados_cli = target_f.evaluate('''() => {
                        const buscar = (t) => {
                            const items = Array.from(document.querySelectorAll('.detail-item'));
                            const found = items.find(i => i.innerText.includes(t));
                            return found ? found.querySelector('.value').innerText.replace("****", "").trim() : "N/A";
                        };
                        return { nome: buscar("Customer Name"), tel: buscar("Customer phone"), end: buscar("Address"), bairro: buscar("Customer District"), cidade: buscar("Recipient City") };
                    }''')

                    try:
                        try: target_f.get_by_role("tab", name="PRODUCT INFO").first.click(force=True)
                        except: target_f.get_by_text("PRODUCT INFO").first.click(force=True)
                        page.wait_for_timeout(1000) 
                        
                        try:
                            olho_prod = target_f.locator('.detail-item', has_text="Goods name").locator('svg, [role="button"]').first
                            if olho_prod.is_visible(timeout=1000):
                                olho_prod.click(force=True)
                                page.wait_for_timeout(1500)
                        except: pass

                        try:
                            olho_val = target_f.locator('.detail-item', has_text="Declared Value (Uploaded)(BRL)").locator('svg, [role="button"]').first
                            if olho_val.is_visible(timeout=1000):
                                olho_val.click(force=True)
                                page.wait_for_timeout(1500)
                        except: pass
                        
                        dados_prod = target_f.evaluate('''() => {
                            const buscar = (t) => {
                                const items = Array.from(document.querySelectorAll('.detail-item'));
                                const found = items.find(i => i.innerText.includes(t));
                                return found ? found.querySelector('.value').innerText.replace("****", "").trim() : "N/A";
                            };
                            return { produto: buscar("Goods name"), valor: buscar("Declared Value (Uploaded)(BRL)") };
                        }''')
                        produto = dados_prod['produto']
                        valor = dados_prod['valor']
                    except: 
                        produto = "N/A"
                        valor = "0.00"

                    prazo = prazo_da_linha if prazo_da_linha else "N/A"
                    try:
                        try: target_f.get_by_role("tab", name="OPS INFO").first.click(force=True)
                        except: target_f.get_by_text("OPS INFO").first.click(force=True)
                        page.wait_for_timeout(1000)
                        
                        motorista = target_f.evaluate('''() => {
                            const items = Array.from(document.querySelectorAll('.detail-item'));
                            const found = items.find(i => i.innerText.includes("DA"));
                            return found ? found.querySelector('.value').innerText.replace("****", "").trim() : "N/A";
                        }''')
                    except: motorista = "N/A"

                    endereco_completo = f"{dados_cli['end']}, {dados_cli['bairro']}, {dados_cli['cidade']}".replace(", N/A", "").replace("N/A, ", "")

                    dados_finais_excel.append({
                        "AWB": codigo,
                        "Motorista": motorista,
                        "Subtipo": sub_tipo_da_linha, 
                        "Nome": dados_cli['nome'],
                        "Telefone": dados_cli['tel'],
                        "Endereco": endereco_completo,
                        "Produto": produto,
                        "Valor": valor,
                        "Prazo do Processo": prazo
                    })

            if dados_finais_excel:
                # ==========================================
                # NOVO: SEPARA OS ARQUIVOS POR PASTA (DRIVES)
                # ==========================================
                pasta_base = f"Arquivos_{sigla}"
                os.makedirs(pasta_base, exist_ok=True) # Cria a pasta da base se não existir
                
                df_final = pd.DataFrame(dados_finais_excel)
                # Salva o arquivo dentro da pasta correspondente
                nome_arquivo = os.path.join(pasta_base, f"dados_acareacoes_{sigla}.xlsx")
                df_final.to_excel(nome_arquivo, index=False)
                
                log(f"Planilha local gerada na pasta '{pasta_base}', subindo para a nuvem...")
                enviar_para_google_sheets(df_final, sigla, bot, chat_id)
                
                # ==========================================
                # MENSAGEM COM O RESUMO DOS MOTORISTAS (LIMPO)
                # ==========================================
                if bot and chat_id:
                    contagem_motoristas = df_final['Motorista'].value_counts()
                    msg_resumo = f"🚚 *MOTORISTAS COM ACAREAÇÕES - {sigla}*\n\n"
                    
                    for mot_completo, qtd in contagem_motoristas.items():
                        nome_limpo = re.sub(r'\(.*?\)', '', str(mot_completo)).strip()
                        nome_limpo = nome_limpo.split('-')[0].strip()
                        primeiro_nome = nome_limpo.split(' ')[0].strip().upper()
                        
                        msg_resumo += f"👤 {primeiro_nome}: {qtd} pacote(s)\n"
                        
                    msg_resumo += f"\n📦 *Total: {len(df_final)} acareação(ões)*"
                    bot.send_message(chat_id, msg_resumo, parse_mode="Markdown")
                
                return nome_arquivo
            else:
                return None

        except Exception as e:
            log(f"❌ Erro crítico: {e}")
            return None
        finally:
            browser.close()