import streamlit as st
import pandas as pd
import os
from google_utils import get_gspread_client

st.set_page_config(layout="wide", page_title="Painel de Acareações", page_icon="🕵️‍♀️")
st.title("🕵️‍♀️ Página de Depuração e Gerador de Mensagens")

# ==========================================
# SEÇÃO 1: DEPURAÇÃO
# ==========================================
with st.expander("Verificar Ambiente (Debug)", expanded=False):
    st.header("1. Onde os arquivos estão?")
    try:
        root_dir = "/mount/src/"
        st.write(f"Listando arquivos em `{root_dir}`:")
        files_in_root = os.listdir(root_dir)
        st.code('\n'.join(files_in_root))
    except Exception as e:
        st.error(f"Erro ao tentar listar arquivos: {e}")

    st.header("2. O 'requirements.txt' está sendo lido corretamente?")
    try:
        req_path = "/mount/src/requirements.txt"
        st.write(f"Tentando ler o arquivo em `{req_path}`...")
        if os.path.exists(req_path):
            with open(req_path, 'r') as f:
                content = f.read()
            st.success("✅ Arquivo 'requirements.txt' encontrado! Conteúdo:")
            st.code(content)
        else:
            st.error("❌ ERRO CRÍTICO: O arquivo 'requirements.txt' NÃO foi encontrado no diretório raiz.")
            st.info("Verifique se o arquivo está na pasta principal do seu projeto no GitHub.")
    except Exception as e:
        st.error(f"Erro ao tentar ler o requirements.txt: {e}")


# ==========================================
# SEÇÃO 2: GERADOR DE MENSAGENS COM CHECKBOX
# ==========================================
st.divider()
st.header("📋 Painel de Acareações e Disparo")

# Opção para o usuário escolher qual base/aba carregar
aba_escolhida = st.text_input("Digite o nome da base (Ex: JFR, BHZ) para carregar os dados:", value="JFR")

if aba_escolhida:
    try:
        # Tenta buscar os dados do Google Sheets
        cliente = get_gspread_client()
        planilha = cliente.open(os.getenv("NOME_PLANILHA", "acareaBase"))
        worksheet = planilha.worksheet(aba_escolhida)
        dados = worksheet.get_all_records()
        df = pd.DataFrame(dados)
    except Exception as e:
        st.warning(f"Não foi possível carregar a aba '{aba_escolhida}' ou ela está vazia. Erro: {e}")
        df = pd.DataFrame()

    if not df.empty:
        # 1. Cria a coluna de Checkbox caso ela não exista
        if "Enviado" not in df.columns:
            df["Enviado"] = False
        else:
            # Garante que o tipo seja booleano
            df["Enviado"] = df["Enviado"].astype(bool)
        
        # 2. Exibe a tabela interativa (Data Editor)
        st.markdown("### Controle de Disparos")
        st.write("Marque a caixa na coluna **Enviado** assim que disparar a mensagem para o motorista.")
        
        df_editado = st.data_editor(
            df,
            column_config={
                "Enviado": st.column_config.CheckboxColumn(
                    "Já Enviado? ✅",
                    help="Marque se a mensagem já foi enviada para o motorista",
                    default=False,
                )
            },
            # Bloqueia a edição das outras colunas para evitar acidentes (ajuste as colunas confome sua base real)
            disabled=["AWB", "Motorista", "Subtipo", "Nome", "Telefone", "Bairro", "Endereco", "Produto", "Valor", "Prazo do Processo"],
            hide_index=True,
            use_container_width=True
        )
        
        # 3. Atualizar Google Sheets com os Checkboxes (Opcional, mas útil se quiser manter salvo)
        if st.button("💾 Salvar status de envio na Nuvem"):
            try:
                worksheet.clear()
                # Salva o dataframe com o cabeçalho
                worksheet.update([df_editado.columns.values.tolist()] + df_editado.values.tolist())
                st.toast("Status salvo com sucesso!")
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")

        st.divider()
        st.subheader("💬 Gerador de Mensagens (Pendentes)")
        
        # 4. Filtra apenas os que ainda NÃO foram marcados como "Enviado"
        pendentes = df_editado[df_editado["Enviado"] == False]
        
        if pendentes.empty:
            st.success("🎉 Todas as mensagens já foram marcadas como enviadas!")
        else:
            # Organiza a exibição das caixas de texto em colunas
            cols = st.columns(2)
            
            for index, row in pendentes.iterrows():
                # Formatação da mensagem incluindo o Bairro e o Telefone
                primeiro_nome_mot = str(row.get('Motorista', 'Motorista')).split(' ')[0]
                
                mensagem = f"""Olá {primeiro_nome_mot}, temos uma acareação pendente para você!

📦 *AWB:* {row.get('AWB', '')}
🛑 *Motivo:* {row.get('Subtipo', 'N/A')}
👤 *Cliente:* {row.get('Nome', '')}
📞 *Telefone:* {row.get('Telefone', 'N/A')}
🏘️ *Bairro:* {row.get('Bairro', 'N/A')}
📍 *Endereço:* {row.get('Endereco', '')}

Por favor, verifique o comprovante e retorne o contato com o cliente!"""
                
                # Distribui nas colunas
                with cols[index % 2]:
                    st.text_area(
                        label=f"Motorista: {primeiro_nome_mot} | AWB: {row.get('AWB', '')}",
                        value=mensagem,
                        height=230,
                        key=f"msg_{row.get('AWB', index)}"
                    )
