import os
import streamlit as st
import pandas as pd
import urllib.parse
import re
import gspread
import json
from google.oauth2.service_account import Credentials
import base64
import requests
from datetime import datetime, timedelta
import plotly.express as px

# ==============================================================
# CONFIGURAÇÃO DA PÁGINA (Deve ser a primeira linha do Streamlit)
# ==============================================================
st.set_page_config(page_title="Gestão de Acareações", layout="wide", page_icon="📦")

# ==============================================================
# CONFIGURAÇÕES E SECRETS
# ==============================================================
try:
    NUMERO_BASE = st.secrets.get("NUMERO_BASE", "5531971463005")
    NUMERO_BASE_CTG = st.secrets.get("NUMERO_BASE_CTG", NUMERO_BASE) 
    NOME_PLANILHA = st.secrets.get("NOME_PLANILHA", "acareaBase")
    URL_WEBHOOK_DRIVE = st.secrets.get("URL_WEBHOOK_DRIVE", "")
except:
    NUMERO_BASE = os.getenv("NUMERO_BASE", "5531971463005")
    NUMERO_BASE_CTG = os.getenv("NUMERO_BASE_CTG", NUMERO_BASE)
    NOME_PLANILHA = os.getenv("NOME_PLANILHA", "acareaBase")
    URL_WEBHOOK_DRIVE = os.getenv("URL_WEBHOOK_DRIVE", "")

# 🔥 CORREÇÃO 1: As 11 bases da rede em ordem alfabética
BASES_DISPONIVEIS = ["CPH", "CTG", "CTP", "GNH", "GVR", "ITR", "JML", "MNT", "QHG", "RBN", "TFO"]

# ==============================================================
# FUNÇÕES DE APOIO E DADOS
# ==============================================================
def obter_credenciais():
    try:
        with open('credenciais.json') as f:
            info = json.load(f)
    except FileNotFoundError:
        info = json.loads(st.secrets["chave_google"].strip())
    escopo = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    return Credentials.from_service_account_info(info, scopes=escopo)

def formatar_prazo(prazo_planilha):
    prazo_texto = str(prazo_planilha).strip()
    if not prazo_texto or prazo_texto.lower() in ['nan', 'none', 'n/a']:
        amanha = datetime.now() + timedelta(days=1)
        return amanha.strftime("%d/%m/%Y") + " 14:00"

    try:
        try: 
            prazo_dt = datetime.strptime(prazo_texto, "%Y-%m-%d %H:%M:%S")
        except: 
            prazo_dt = datetime.strptime(prazo_texto[:16], "%Y-%m-%d %H:%M")

        prazo_ajustado = prazo_dt - timedelta(hours=2)
        hora = prazo_ajustado.hour

        if hora >= 16 or hora < 6:
            return prazo_ajustado.strftime("%d/%m/%Y") + " 14:00"
        else:
            return prazo_ajustado.strftime("%d/%m/%Y %H:%M")
    except:
        return datetime.now().strftime("%d/%m/%Y") + " 14:00"

def upload_para_drive(arquivo_foto, nome_arquivo):
    if not URL_WEBHOOK_DRIVE:
        st.error("URL do Webhook do Google Apps Script não configurada.")
        return None
    try:
        bytes_imagem = arquivo_foto.getvalue()
        base64_img = base64.b64encode(bytes_imagem).decode('utf-8')

        payload = {
            "filename": nome_arquivo,
            "mimetype": arquivo_foto.type,
            "base64": base64_img
        }

        resposta = requests.post(URL_WEBHOOK_DRIVE, json=payload)
        resultado = resposta.json()

        if resultado.get("status") == "success":
            return resultado.get("id")
        else:
            st.error(f"Erro no Drive: {resultado.get('message')}")
            return None
    except Exception as e:
        st.error(f"Erro de conexão com o Drive Webhook: {e}")
        return None

@st.cache_data(ttl=60)
def carregar_dados_base(nome_aba):
    try:
        creds = obter_credenciais()
        cliente = gspread.authorize(creds)
        planilha = cliente.open(NOME_PLANILHA).worksheet(nome_aba)
        dados = planilha.get_all_values()

        if not dados or len(dados) < 2: 
            return pd.DataFrame()

        df = pd.DataFrame(dados[1:], columns=dados[0])
        df.columns = df.columns.astype(str).str.strip()
        if 'Motorista' in df.columns:
            df['Motorista'] = df['Motorista'].astype(str).str.strip()
        return df
    except: 
        return pd.DataFrame()

@st.cache_data(ttl=120)
def carregar_kpi_historico():
    try:
        creds = obter_credenciais()
        cliente = gspread.authorize(creds)
        planilha = cliente.open(NOME_PLANILHA).worksheet("KPI_Historico")
        dados = planilha.get_all_values()
        if len(dados) > 1:
            df = pd.DataFrame(dados[1:], columns=dados[0])
            df['KPI (%)'] = df['KPI (%)'].astype(str).str.replace('%', '').astype(float)
            df['Data'] = pd.to_datetime(df['Data'])
            df['Entregues'] = pd.to_numeric(df['Entregues'])
            df['Acareações'] = pd.to_numeric(df['Acareações'])
            return df
        return pd.DataFrame()
    except Exception as e:
        return pd.DataFrame()

# ==============================================================
# MENU LATERAL (SIDEBAR)
# ==============================================================
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/ca/Imile_Delivery_Logo.jpg/800px-Imile_Delivery_Logo.jpg", width=150)
st.sidebar.title("Navegação")

# Corrigido menu duplicado
menu = st.sidebar.radio("Ir para:", ["📷 Portal do Motorista", "🚨 Painel de Risco (< 5h)", "🔥 Fila de Urgências", "📈 Dashboard de KPI", "💬 Gerador de Mensagens"])
query_params = st.query_params
base_via_url = query_params.get("base", None)

# ==============================================================
# TELA 1: PORTAL DO MOTORISTA
# ==============================================================
if menu == "📷 Portal do Motorista":
    st.title("📦 Portal de Acareações")

    if base_via_url in BASES_DISPONIVEIS:
        base_atual = base_via_url
        st.info(f"📍 Base detectada: **{base_atual}**")
    else:
        base_atual = st.selectbox("🏢 Selecione sua Unidade:", ["-- Escolha --"] + BASES_DISPONIVEIS)

    if base_atual != "-- Escolha --":
        df_imile = carregar_dados_base(base_atual)
        if not df_imile.empty:
            motoristas = sorted(df_imile['Motorista'].dropna().unique().tolist())
            if '(vazio)' in motoristas: motoristas.remove('(vazio)')
            mot_selecionado = st.selectbox("👤 Selecione o seu nome:", ["-- Escolha --"] + motoristas)

            if mot_selecionado != "-- Escolha --":
                df_mot = df_imile[df_imile['Motorista'] == mot_selecionado].copy()
                st.success(f"Você tem **{len(df_mot)} acareação(ões)** pendente(s) na base {base_atual}.")

                if 'Subtipo' not in df_mot.columns:
                    df_mot['Subtipo'] = "N/A"
                df_mot['Subtipo'] = df_mot['Subtipo'].apply(lambda x: "N/A" if pd.isna(x) or str(x).strip() == "" else str(x).strip())

                subtipos_unicos = sorted(df_mot['Subtipo'].unique())

                for subtipo_atual in subtipos_unicos:
                    df_sub = df_mot[df_mot['Subtipo'] == subtipo_atual]
                    st.markdown(f"<h3 style='color: #0068C9; margin-top: 25px; border-bottom: 2px solid #0068C9; padding-bottom: 5px;'>🏷️ {subtipo_atual} <span style='color: gray; font-size: 16px;'>({len(df_sub)} pacote(s))</span></h3>", unsafe_allow_html=True)

                    for idx, row in df_sub.iterrows():
                        with st.expander(f"🔵 iMile | Pacote: {row['AWB']} - {row['Nome']}", expanded=False):
                            prazo_texto = formatar_prazo(row.get('Prazo do Processo', ''))

                            st.info(f"⏰ PRAZO DE FECHAMENTO: {prazo_texto}")
                            st.info(f"💰 VALOR DO PACOTE: RS {row.get('Valor', '0.00')} + R$ 100,00 MULTA")

                            st.markdown("### 📷 Enviar Comprovante")
                            foto = st.file_uploader(f"Anexe o print/foto (AWB {row['AWB']})", type=['png', 'jpg', 'jpeg'], key=f"file_{row['AWB']}")

                            if foto:
                                if st.button(f"Confirmar Envio da Foto {row['AWB']}", key=f"btn_{row['AWB']}"):
                                    with st.spinner("Enviando para a base..."):
                                        nome_img = f"{base_atual}_{row['AWB']}_{row['Nome']}.jpeg".replace(" ", "_")
                                        file_id = upload_para_drive(foto, nome_img)
                                        if file_id:
                                            st.success("✅ Foto salva com sucesso no Google Drive!")
                                            st.balloons()

                            st.markdown("---")
                            tel_bruto = str(row.get('Telefone', ''))
                            tel_cliente = re.sub(r'\D', '', tel_bruto).lstrip('0')
                            if len(tel_cliente) >= 10: tel_cliente = '55' + tel_cliente
                            else: tel_cliente = ''

                            msg_cliente = (
                                f"Olá, somos uma transportadora parceira (SHEIN/TIKTOK)\n\n"
                                f"{row['Nome']}, poderia confirmar o recebimento da mercadoria com os dados abaixo:\n"
                                f"Código do pacote: {row['AWB']}\n"
                                f"Endereço: {row.get('Endereco', 'N/A')}\n\n"
                                f"Produto: {row.get('Produto', 'N/A')}\n\n"
                                f"Confirma o Recebimento do produto? SIM OU NÃO"
                            )
                            st.code(msg_cliente, language="text") 

                            col1, col2 = st.columns(2)
                            with col1:
                                if tel_cliente:
                                    st.link_button("1️⃣ Chamar Cliente", f"https://wa.me/{tel_cliente}?text={urllib.parse.quote(msg_cliente)}")
                                else:
                                    st.error("Telefone indisponível")
                            with col2:
                                msg_base = f"Base, segue comprovante do pacote {row['AWB']}."
                                numero_whats = NUMERO_BASE_CTG if base_atual == "CTG" else NUMERO_BASE
                                st.link_button("2️⃣ Avisar Base", f"https://wa.me/{numero_whats}?text={urllib.parse.quote(msg_base)}")
        else:
            st.warning(f"⚠️ Nenhuma acareação pendente na base {base_atual}.")

# ==============================================================
# TELA 2: PAINEL DE RISCO (< 5 HORAS)
# ==============================================================
elif menu == "🚨 Painel de Risco (< 5h)":
    st.title("🚨 Acareações em Risco de Vencimento")
    st.markdown("Pacotes que vencem nas próximas **5 horas** e precisam de atenção urgente.")

    agora = datetime.now()
    limite_5h = agora + timedelta(hours=5)

    lista_risco = []

    with st.spinner("Analisando os prazos de todas as bases..."):
        for base in BASES_DISPONIVEIS:
            df_base = carregar_dados_base(base)
            if not df_base.empty:
                for idx, row in df_base.iterrows():
                    prazo_texto = str(row.get('Prazo do Processo', '')).strip()
                    if prazo_texto and prazo_texto.lower() not in ['nan', 'none', 'n/a']:
                        try:
                            try: prazo_dt = datetime.strptime(prazo_texto, "%Y-%m-%d %H:%M:%S")
                            except: prazo_dt = datetime.strptime(prazo_texto[:16], "%Y-%m-%d %H:%M")

                            if prazo_dt <= limite_5h:
                                lista_risco.append({
                                    "Base": base,
                                    "Motorista": row.get('Motorista', 'N/A'),
                                    "AWB": row['AWB'],
                                    "Motivo": row.get('Subtipo', 'N/A'),
                                    "Prazo Limite": prazo_dt.strftime("%d/%m %H:%M")
                                })
                        except: pass

    if lista_risco:
        df_risco = pd.DataFrame(lista_risco)
        st.error(f"⚠️ Existem {len(df_risco)} acareações estourando o prazo!")
        st.dataframe(df_risco, use_container_width=True)
    else:
        st.success("✅ Tudo sob controle! Nenhuma acareação próxima do vencimento nas próximas 5 horas nas bases ativas.")

# ==============================================================
# 🔹 TELA 3: NOVO PAINEL DE URGÊNCIAS (LATÊNCIA ALTA)
# ==============================================================
elif menu == "🔥 Fila de Urgências":
    st.title("🔥 Fila de Urgências de Inventário")
    st.markdown("Acompanhe os pacotes críticos com **Latência ≥ 3 dias** parados na base ou na mão de motoristas. Cobre a equipe para limpar essa fila!")

    with st.spinner("Buscando dados críticos..."):
        df_urgencias = carregar_dados_base("Urgencias_Latencia")

    if df_urgencias.empty:
        st.success("✅ **EXCELENTE!** Não há pacotes com latência crítica (≥ 3 dias) em nenhuma base no momento.")
        st.balloons()
    else:
        if 'Dias Parado' in df_urgencias.columns:
            df_urgencias['Dias Parado'] = pd.to_numeric(df_urgencias['Dias Parado'], errors='coerce')

        col1, col2 = st.columns(2)
        with col1:
            bases_urg = ["Todas"] + list(df_urgencias['Base'].dropna().unique())
            filtro_base = st.selectbox("Filtrar por Base:", bases_urg)

        with col2:
            mots_urg = ["Todos"] + sorted([str(m) for m in df_urgencias['Motorista'].dropna().unique() if str(m).strip() != ''])
            filtro_mot = st.selectbox("Filtrar por Motorista:", mots_urg)

        df_show = df_urgencias.copy()
        if filtro_base != "Todas":
            df_show = df_show[df_show['Base'] == filtro_base]
        if filtro_mot != "Todos":
            df_show = df_show[df_show['Motorista'] == filtro_mot]

        st.warning(f"⚠️ **Total de pacotes críticos encontrados:** {len(df_show)}")

        st.dataframe(
            df_show,
            use_container_width=True,
            hide_index=True
        )

# ==============================================================
# TELA 4: DASHBOARD DE KPI (GRÁFICOS)
# ==============================================================
elif menu == "📈 Dashboard de KPI":
    st.title("📈 Análise de Acareações vs Entregas")

    df_kpi = carregar_kpi_historico()

    if df_kpi.empty:
        st.warning("⚠️ O histórico ainda está vazio. O robô precisa rodar o arquivo 'kpi_mensal.py' para gerar os primeiros dados.")
    else:
        filtro_tempo = st.radio("Filtro de Período:", ["Últimos 7 dias", "Últimos 15 dias", "Mês Atual", "Tudo"], horizontal=True)

        hoje = pd.to_datetime(datetime.now().date())
        if filtro_tempo == "Últimos 7 dias": df_kpi = df_kpi[df_kpi['Data'] >= (hoje - pd.Timedelta(days=7))]
        elif filtro_tempo == "Últimos 15 dias": df_kpi = df_kpi[df_kpi['Data'] >= (hoje - pd.Timedelta(days=15))]
        elif filtro_tempo == "Mês Atual": df_kpi = df_kpi[df_kpi['Data'].dt.month == hoje.month]

        df_kpi = df_kpi.drop_duplicates(subset=['Data', 'Base'], keep='last')
        df_kpi = df_kpi.sort_values(by='Data')
        df_kpi['Data_Formatada'] = df_kpi['Data'].dt.strftime('%d/%m')

        st.markdown("---")

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📊 Taxa de Acareação % (KPI)")
            fig_kpi = px.line(
                df_kpi, 
                x='Data_Formatada', 
                y='KPI (%)', 
                color='Base', 
                labels={'Data_Formatada': 'Data', 'KPI (%)': 'KPI (%)'}
            )
            fig_kpi.update_traces(mode='lines+markers')
            fig_kpi.update_layout(xaxis_title="", yaxis_title="KPI (%)", legend_title="Base", hovermode="x unified")
            st.plotly_chart(fig_kpi, use_container_width=True)

        with col2:
            st.subheader("📦 Total de Acareações Geradas")
            fig_qnt = px.bar(
                df_kpi, 
                x='Data_Formatada', 
                y='Acareações', 
                color='Base', 
                barmode='group',
                text_auto=True,
                labels={'Data_Formatada': 'Data', 'Acareações': 'Qtd Acareações'}
            )
            fig_qnt.update_layout(xaxis_title="", yaxis_title="Acareações", legend_title="Base", hovermode="x unified")
            st.plotly_chart(fig_qnt, use_container_width=True)

        st.markdown("---")
        st.subheader("📋 Tabela Consolidada (Período Selecionado)")
        resumo = df_kpi.groupby('Base').agg({'Entregues': 'sum', 'Acareações': 'sum'}).reset_index()
        resumo['KPI Geral (%)'] = resumo.apply(
            lambda row: round((row['Acareações'] / row['Entregues']) * 100, 2) if row['Entregues'] > 0 else 0.0, 
            axis=1
        )
        
        # 🔥 CORREÇÃO 2: Ordenar a tabela de Resumo em Ordem Alfabética
        resumo = resumo.sort_values(by='Base').reset_index(drop=True)
        
        st.dataframe(resumo, use_container_width=True)

# ==============================================================
# TELA 5: GERADOR DE MENSAGENS PARA O MOTORISTA (ANALISTAS)
# ==============================================================
elif menu == "💬 Gerador de Mensagens":
    st.title("💬 Gerador de Mensagens de Acareação")
    st.markdown("Gere o texto padrão de cobrança com todas as regras (multa, print, protocolo) para enviar aos motoristas no WhatsApp.")

    base_atual = st.selectbox("🏢 Selecione a Base:", ["-- Escolha --"] + BASES_DISPONIVEIS)

    if base_atual != "-- Escolha --":
        df_imile = carregar_dados_base(base_atual)
        if not df_imile.empty:
            motoristas = sorted(df_imile['Motorista'].dropna().unique().tolist())
            if '(vazio)' in motoristas: motoristas.remove('(vazio)')
            mot_selecionado = st.selectbox("👤 Selecione o Motorista:", ["-- Escolha --"] + motoristas)

            if mot_selecionado != "-- Escolha --":
                df_mot = df_imile[df_imile['Motorista'] == mot_selecionado].copy()
                
                awbs = df_mot['AWB'].unique().tolist()
                awb_selecionado = st.selectbox("📦 Selecione o Pacote (AWB):", ["-- Escolha --"] + awbs)
                
                if awb_selecionado != "-- Escolha --":
                    linha = df_mot[df_mot['AWB'] == awb_selecionado].iloc[0]
                    
                    prazo_texto = formatar_prazo(linha.get('Prazo do Processo', ''))
                    produto = str(linha.get('Produto', 'N/A')).strip()
                    telefone = str(linha.get('Telefone', 'N/A')).strip()
                    nome_cliente = str(linha.get('Nome', 'Cliente')).strip()
                    
                    texto_gerado = f"""Boa tarde! Segue acareação iMile:

Código do pedido: {awb_selecionado}
Produto: {produto}
Telefone: {telefone}

Finalizar até: {prazo_texto}

Você pode responder as acareações com um print com a conversa com o cliente confirmando o recebimento (incluindo código do pedido, nome e telefone do mesmo, o contato não pode estar salvo e o telefone tem que ser o mesmo do sistema) ou protocolo de acareação por escrito preenchendo todos os dados corretamente (tenho impresso, pegar na base). O cabeçalho NÃO pode ser preenchido pelo cliente, pra não ficar tudo com a mesma letra, e uma foto da fachada da casa. 

Caso não seja respondido dentro do prazo, ela irá extraviar em seu nome + multa de R$100,00 reais por baixa fake.
(CASO NÃO SEJA RESPONDIDA NO PRAZO DA FORMA CORRETA, SERÁ DESCONTADO)

Exemplo de como mandar mensagem:

Oi, boa tarde {nome_cliente}, sou o entregador, gostaria de confirmar o recebimento de sua mercadoria com o código {awb_selecionado}.
Produto: {produto}"""

                    st.markdown("### 📋 Texto Pronto para Copiar")
                    st.code(texto_gerado, language="text")
                    
                    st.markdown("---")
                    
                    msg_url = urllib.parse.quote(texto_gerado)
                    st.link_button("📲 Enviar direto via WhatsApp", f"https://wa.me/?text={msg_url}")
                    
        else:
            st.warning(f"⚠️ Nenhuma acareação pendente na base {base_atual}.")