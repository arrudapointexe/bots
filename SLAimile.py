import telebot
import subprocess
import sys
import os
import logging
from playwright.sync_api import sync_playwright
import re
from dotenv import load_dotenv
import schedule
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configuração para mostrar apenas ERROS no terminal
logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

load_dotenv()

# Importações dos seus módulos internos
from gerar_perdas import extrair_relatorio_diario_completo
from SLAimile import baixar_inventario, gerar_prints_e_mensagens
from acareacoes import rodar_automacao_acareacao, rodar_automacao_acareacao_incremental
from shopee import carregar_dicionario_ceps, baixar_planilha_shopee, gerar_relatorio_shopee
from config import CHAT_ID_ALVO, BASES_ACAREACOES, BASES_SLA

TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
bot = telebot.TeleBot(TOKEN)

# ==============================================================
# FLUXO 1: RELATÓRIOS DE SLA / LATÊNCIA
# ==============================================================
def executar_fluxo_sla(sigla, usuario, senha, chat_id):
    # Removidos os avisos intermediários para limpar os logs
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False) 
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            page.goto("https://ds-login.imile.com/", wait_until="domcontentloaded")
            
            page.fill('input[type="text"]', str(usuario or ""))
            page.fill('input[type="password"]', str(senha or ""))
            page.get_by_role("button", name=re.compile(r"登录|Conecte-se|Login", re.IGNORECASE)).click()
            page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(5000)
            
            try: page.locator('.close-icon').first.click(force=True)
            except: pass

            arquivo_bruto = baixar_inventario(page, sigla)
            
            if arquivo_bruto and os.path.exists(arquivo_bruto):
                dados_telegram = gerar_prints_e_mensagens(arquivo_bruto, sigla, page)

                if dados_telegram:
                    try:
                        with open(dados_telegram["img_lat"], 'rb') as f_lat:
                            bot.send_photo(chat_id, f_lat)
                        bot.send_message(chat_id, dados_telegram["txt_lat"], parse_mode="Markdown")
                    except Exception as e:
                        bot.send_message(chat_id, f"⚠️ Erro ao enviar Latência: {e}")

                    for img_rota, txt_rota in zip(dados_telegram["imgs_rota"], dados_telegram["txts_rota"]):
                        try:
                            with open(img_rota, 'rb') as f_rota:
                                bot.send_photo(chat_id, f_rota, caption=txt_rota, parse_mode="Markdown")
                        except Exception as e:
                            bot.send_message(chat_id, f"⚠️ Erro ao enviar Rota: {e}")

                    bot.send_message(chat_id, f"✅ Processo SLA {sigla} finalizado com sucesso!")
                else:
                    bot.send_message(chat_id, f"❌ Erro: Falha ao gerar as imagens para {sigla}.")
            else:
                bot.send_message(chat_id, f"❌ Erro: Arquivo bruto não foi gerado para {sigla}.")

        except Exception as e:
            bot.send_message(chat_id, f"❌ Erro crítico na base {sigla}: {e}")
            logging.error(f"Erro no fluxo SLA para {sigla}: {e}")
        finally:
            browser.close()

# ==============================================================
# FLUXO 2: AUTOMAÇÃO DE ACAREAÇÕES E CALCULO DE RESUMO
# ==============================================================
def executar_fluxo_acareacao(sigla, usuario, senha, chat_id):
    arquivo_gerado = rodar_automacao_acareacao(sigla, usuario, senha, bot, chat_id)
    
    if arquivo_gerado and os.path.exists(arquivo_gerado):
        try:
            with open(arquivo_gerado, 'rb') as doc:
                bot.send_document(chat_id, doc, caption=f"📦 Planilha de Acareações {sigla} extraída e enviada ao Drive com sucesso!")
        except Exception as e:
            bot.send_message(chat_id, f"⚠️ A automação rodou, mas falhou ao enviar o Excel aqui no Telegram: {e}")
    else:
        bot.send_message(chat_id, f"✅ Processo de acareação da {sigla} finalizado, mas nenhum dado novo precisou ser salvo.")


def executar_fluxo_acareacao_incremental(sigla, usuario, senha, chat_id):
    arquivo_gerado = rodar_automacao_acareacao_incremental(sigla, usuario, senha, bot, chat_id)
    
    if arquivo_gerado and os.path.exists(arquivo_gerado):
        try:
            with open(arquivo_gerado, 'rb') as doc:
                bot.send_document(chat_id, doc, caption=f"📦 Planilha de Acareações {sigla} extraída e enviada ao Drive com sucesso!")
        except Exception as e:
            bot.send_message(chat_id, f"⚠️ A automação rodou, mas falhou ao enviar o Excel aqui no Telegram: {e}")
    else:
        bot.send_message(chat_id, f"✅ Processo de acareação da {sigla} finalizado, mas nenhum dado novo precisou ser salvo.")

def calcular_resumo_base(caminho_excel):
    from datetime import datetime, timedelta
    import pandas as pd
    if not caminho_excel or not os.path.exists(caminho_excel):
        return 0, 0
    try:
        df = pd.read_excel(caminho_excel)
        total = len(df)
        urgentes = 0
        agora = datetime.now()
        limite_5h = agora + timedelta(hours=5)
        
        if 'Prazo do Processo' in df.columns:
            for prazo_texto in df['Prazo do Processo'].dropna().astype(str):
                try:
                    try:
                        prazo_dt = datetime.strptime(prazo_texto.strip(), "%Y-%m-%d %H:%M:%S")
                    except:
                        prazo_dt = datetime.strptime(prazo_texto.strip()[:16], "%Y-%m-%d %H:%M")
                    
                    if prazo_dt <= limite_5h:
                        urgentes += 1
                except: pass
        return total, urgentes
    except Exception as e:
        logging.error(f"Erro ao calcular resumo: {e}")
        return 0, 0

# ==============================================================
# COMANDOS PRINCIPAIS DO BOT
# ==============================================================
@bot.message_handler(commands=['ajuda', 'start'])
def send_welcome(message):
    texto = (
        "Olá! Escolha um comando abaixo:\n\n"
        "🚚 *SHOPEE:*\n"
        "/shopee - Gerar Tabela Dinâmica de Vencimentos\n\n"
        "📊 *Relatórios:*\n"
        "/kpi - Atualiza o Dashboard Mensal\n"
        "/perdas - Relatório diário de inventários\n\n"
        "📈 *SLA e Latência:*\n"
        "/jml - Extrair Latência JML\n"
        "/itr - Extrair Latência ITR\n"
        "/ctg - Extrair Latência CTG\n\n"
        "📦 *Automação de Acareações:*\n"
        "/acarea - Rodar todas as bases no Modo Turbo\n"
        "/acareaatualitr - Atualizar só os novos itens na planilha\n"
        "/acareajml - Buscar e subir acareações JML\n"
        "/acareaitr - Buscar e subir acareações ITR\n"
        "/acareactg - Buscar e subir acareações CTG\n"
    )
    bot.reply_to(message, texto, parse_mode="Markdown")

@bot.message_handler(commands=['shopee'])
def command_shopee(message):
    bot.send_message(message.chat.id, "🔍 Iniciando a extração do Export Forward no portal da Shopee. Aguarde...")
    caminho_forward = baixar_planilha_shopee()
    if not caminho_forward or not os.path.exists(caminho_forward):
        bot.send_message(message.chat.id, "❌ Erro ao baixar o arquivo da Shopee. Verifique o portal ou o login.")
        return

    caminho_ceps = "base_ceps.xlsx"
    mapa_ceps = carregar_dicionario_ceps(caminho_ceps)
    if not mapa_ceps:
        bot.send_message(message.chat.id, "⚠️ Aviso: Base de CEPs vazia. Alguns pacotes ficarão sem cidade.")

    bot.send_message(message.chat.id, "📊 Planilha baixada! Gerando as imagens por rota...")
    mensagens, imagens = gerar_relatorio_shopee(caminho_forward, mapa_ceps)
    
    if not imagens: 
        for msg in mensagens:
            bot.send_message(message.chat.id, msg)
    else:
        for i, img in enumerate(imagens):
            if os.path.exists(img):
                with open(img, 'rb') as foto:
                    bot.send_photo(message.chat.id, foto, caption=mensagens[i], parse_mode="Markdown")
        if len(mensagens) > len(imagens):
            for aviso_extra in mensagens[len(imagens):]:
                bot.send_message(message.chat.id, aviso_extra)

@bot.message_handler(commands=['meuid'])
def descobrir_id(message):
    bot.reply_to(message, f"O ID deste chat é: `{message.chat.id}`", parse_mode="Markdown")

@bot.message_handler(commands=['kpi'])
def extrair_kpi_comando(message):
    chat_id = message.chat.id
    bot.send_message(chat_id, "📊 *Iniciando Extração de KPI!*\nO robô está varrendo a iMile para calcular as métricas. Isso leva alguns minutos...", parse_mode="Markdown")
    try:
        caminho_python = sys.executable
        caminho_script = "c:/bots/kpi_mensal.py"
        subprocess.run([caminho_python, caminho_script], check=True)
        bot.send_message(chat_id, "✅ *KPI Atualizado com Sucesso!*\nOs dados já estão disponíveis no seu Dashboard no site.", parse_mode="Markdown")
    except Exception as e:
        bot.send_message(chat_id, f"❌ Ocorreu um erro ao extrair o KPI: {e}")
        logging.error(f"Erro extraindo KPI: {e}")

@bot.message_handler(commands=['perdas'])
def command_perdas(message):
    bot.send_message(message.chat.id, "🔍 A iniciar a extração diária. A descarregar novos inventários e gerar tabelas...")
    try:
        user_jml = os.getenv("IMILE_JML_USER")
        user_itr = os.getenv("IMILE_ITR_USER")
        pass_geral = os.getenv("IMILE_PASS")
        
        if not user_jml or not user_itr or not pass_geral:
            bot.send_message(message.chat.id, "⚠️ Credenciais de JML ou ITR ausentes no .env.")
            return

        msg_perdas, img_itr, img_jml = extrair_relatorio_diario_completo(user_jml, pass_geral, user_itr, pass_geral)
        
        if img_itr and os.path.exists(img_itr):
            with open(img_itr, 'rb') as f_itr:
                bot.send_photo(message.chat.id, f_itr)
        if img_jml and os.path.exists(img_jml):
            with open(img_jml, 'rb') as f_jml:
                bot.send_photo(message.chat.id, f_jml)
                
        bot.send_message(message.chat.id, msg_perdas)
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ocorreu um erro na extração completa: {e}")
        logging.error(f"Erro em perdas: {e}")

# ----------------- COMANDOS SLA (LATÊNCIA) -----------------
@bot.message_handler(commands=['ctg'])
def command_sla_ctg(message):
    user = os.getenv("IMILE_CTG_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_sla("CTG", user, pw, message.chat.id)

@bot.message_handler(commands=['jml'])
def command_sla_jml(message):
    user = os.getenv("IMILE_JML_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_sla("JML", user, pw, message.chat.id)

@bot.message_handler(commands=['itr'])
def command_sla_itr(message):
    user = os.getenv("IMILE_ITR_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_sla("ITR", user, pw, message.chat.id)

# ----------------- COMANDOS DE ACAREAÇÕES MANUAIS -----------------
@bot.message_handler(commands=['acareaatualitr'])
def command_acarea_atual_itr(message):
    user = os.getenv("IMILE_ITR_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao_incremental("ITR", user, pw, message.chat.id)

@bot.message_handler(commands=['acareaatualjml'])
def command_acarea_atual_jml(message):
    user = os.getenv("IMILE_JML_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao_incremental("JML", user, pw, message.chat.id)

@bot.message_handler(commands=['acareactg'])
def command_acarea_ctg(message):
    user = os.getenv("IMILE_CTG_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao("CTG", user, pw, message.chat.id)

@bot.message_handler(commands=['acareajml'])
def command_acarea_jml(message):
    user = os.getenv("IMILE_JML_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao("JML", user, pw, message.chat.id)

@bot.message_handler(commands=['acareaitr'])
def command_acarea_itr(message):
    user = os.getenv("IMILE_ITR_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao("ITR", user, pw, message.chat.id)

@bot.message_handler(commands=['acareagnh'])
def command_acarea_gnh(message):
    user = os.getenv("IMILE_GNH_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao("GNH", user, pw, message.chat.id)

@bot.message_handler(commands=['acareamnt'])
def command_acarea_mnt(message):
    user = os.getenv("IMILE_MNT_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao("MNT", user, pw, message.chat.id)

@bot.message_handler(commands=['acareagvr'])
def command_acarea_gvr(message):
    user = os.getenv("IMILE_GVR_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao("GVR", user, pw, message.chat.id)

@bot.message_handler(commands=['acareaqhg'])
def command_acarea_qhg(message):
    user = os.getenv("IMILE_QHG_USER")
    pw = os.getenv("IMILE_PASS")
    executar_fluxo_acareacao("QHG", user, pw, message.chat.id)

@bot.message_handler(commands=['acarea'])
def acarea_todas(message):
    chat_id = message.chat.id
    bot.send_message(chat_id, "🚀 *A iniciar varredura geral (MODO TURBO)!*\nBuscando dados em segundo plano...", parse_mode="Markdown")

    bases_para_rodar = BASES_ACAREACOES
    resumos = {}

    def processar_base_manual(sigla, usuario, senha):
        if not usuario or not senha:
            logging.error(f"Credenciais da base {sigla} ausentes no .env.")
            return sigla, None
        try:
            caminho_file = rodar_automacao_acareacao(sigla, usuario, senha, bot, chat_id)
            return sigla, caminho_file
        except Exception as e:
            bot.send_message(chat_id, f"❌ Erro ao processar a base {sigla}: {e}")
            logging.error(f"Erro em acarea_todas na base {sigla}: {e}")
            return sigla, None

    with ThreadPoolExecutor(max_workers=5) as executor:
        futuros = [executor.submit(processar_base_manual, b[0], b[1], b[2]) for b in bases_para_rodar]
        for futuro in as_completed(futuros):
            sigla, caminho_file = futuro.result()
            if caminho_file:
                total, urgentes = calcular_resumo_base(caminho_file)
                resumos[sigla] = {"total": total, "urgentes": urgentes}
            else:
                resumos[sigla] = {"total": 0, "urgentes": 0}

    # Ordenar por maior quantidade de acareações totais
    itens_ordenados = sorted(resumos.items(), key=lambda item: item[1]['total'], reverse=True)

    msg_analistas = "📊 *CONSOLIDADO DE ACAREAÇÕES POR BASE* 📊\n━━━━━━━━━━━━━━━━━━━━━━\nOlá equipa, segue o balanço pendente:\n\n"
    total_rede = 0
    
    for sigla, dados in itens_ordenados:
        tot = dados["total"]
        urg = dados["urgentes"]
        total_rede += tot
        txt_urg = f"_(🚨 {urg} CRÍTICOS < 5H)_" if urg > 0 else "_(✅ Sem urgências)_"
        msg_analistas += f"🏢 *{sigla}:* {tot} pacote(s) {txt_urg}\n"
        
    msg_analistas += f"━━━━━━━━━━━━━━━━━━━━━━\n📦 *TOTAL DA REDE:* {total_rede} acareações pendentes."
    bot.send_message(chat_id, msg_analistas, parse_mode="Markdown")

# ==============================================================
# DESPERTADORES INTERNOS
# ==============================================================

def rotina_automatica_acareacoes():
    bot.send_message(CHAT_ID_ALVO, "⏰ *HORÁRIO ATINGIDO! Iniciando varredura automática (Modo Turbo)*", parse_mode="Markdown")
    
    bases_para_rodar = BASES_ACAREACOES
    resumos = {}
    
    def processar_base_auto(sigla, usuario, senha):
        if not usuario or not senha: return sigla, None
        try:
            caminho_file = rodar_automacao_acareacao(sigla, usuario, senha, bot, CHAT_ID_ALVO)
            return sigla, caminho_file
        except Exception as e:
            logging.error(f"Erro em rotina_automatica_acareacoes na base {sigla}: {e}")
            return sigla, None

    with ThreadPoolExecutor(max_workers=5) as executor:
        futuros = [executor.submit(processar_base_auto, b[0], b[1], b[2]) for b in bases_para_rodar]
        for futuro in as_completed(futuros):
            sigla, caminho_file = futuro.result()
            if caminho_file:
                total, urgentes = calcular_resumo_base(caminho_file)
                resumos[sigla] = {"total": total, "urgentes": urgentes}
            else:
                resumos[sigla] = {"total": 0, "urgentes": 0}

    # Ordenar por maior quantidade de acareações totais
    itens_ordenados = sorted(resumos.items(), key=lambda item: item[1]['total'], reverse=True)

    msg_analistas = "📊 *CONSOLIDADO (AUTOMÁTICO)* 📊\n━━━━━━━━━━━━━━━━━━━━━━\nBalanço pendente:\n\n"
    total_rede = 0
    
    for sigla, dados in itens_ordenados:
        tot = dados["total"]
        urg = dados["urgentes"]
        total_rede += tot
        txt_urg = f"_(🚨 {urg} CRÍTICOS < 5H)_" if urg > 0 else "_(✅ Sem urgências)_"
        msg_analistas += f"🏢 *{sigla}:* {tot} pacote(s) {txt_urg}\n"
        
    msg_analistas += f"━━━━━━━━━━━━━━━━━━━━━━\n📦 *TOTAL DA REDE:* {total_rede} acareações."
    bot.send_message(CHAT_ID_ALVO, msg_analistas, parse_mode="Markdown")

def rotina_automatica_sla():
    bot.send_message(CHAT_ID_ALVO, "⏰ *HORÁRIO ATINGIDO! Extração de SLA (Latência)...*", parse_mode="Markdown")
    
    bases_sla = BASES_SLA

    def processar_sla_auto(sigla, usuario, senha):
        if not usuario or not senha: return
        try: executar_fluxo_sla(sigla, usuario, senha, CHAT_ID_ALVO)
        except Exception as e: 
            bot.send_message(CHAT_ID_ALVO, f"❌ Erro SLA {sigla}: {e}")
            logging.error(f"Erro SLA auto {sigla}: {e}")

    with ThreadPoolExecutor(max_workers=5) as executor:
        futuros = [executor.submit(processar_sla_auto, b[0], b[1], b[2]) for b in bases_sla]
        for futuro in as_completed(futuros): pass

def relogio_do_bot():
    while True:
        schedule.run_pending()
        time.sleep(10)

# Configurando os horários fixos
#schedule.every().day.at("08:00").do(rotina_automatica_acareacoes)
#schedule.every().day.at("10:00").do(rotina_automatica_acareacoes)
#schedule.every().day.at("12:00").do(rotina_automatica_acareacoes)
#schedule.every().day.at("14:00").do(rotina_automatica_acareacoes)
#schedule.every().day.at("16:00").do(rotina_automatica_acareacoes)
#schedule.every().day.at("18:00").do(rotina_automatica_acareacoes)
#schedule.every().day.at("20:00").do(rotina_automatica_acareacoes)

schedule.every().day.at("11:50").do(rotina_automatica_sla)
schedule.every().day.at("14:50").do(rotina_automatica_sla)
schedule.every().day.at("17:50").do(rotina_automatica_sla)
schedule.every().day.at("20:50").do(rotina_automatica_sla)

# ==============================================================
# INICIALIZAÇÃO (SEMPRE AQUI NO FINAL)
# ==============================================================
threading.Thread(target=relogio_do_bot, daemon=True).start()
logging.error("⏰ Bot iniciado! (Mostrando apenas erros a partir de agora)")

bot.infinity_polling()
