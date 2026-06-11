
import streamlit as st
import database

# --- Configuração da Página Principal ---
st.set_page_config(
    page_title="Sistema de Bipagem O Boticário",
    page_icon="📦",
    layout="wide"
)

# --- Inicialização ---
# Garante que o banco de dados e a tabela sejam criados na primeira execução
database.setup_database()

# --- Conteúdo da Página Principal ---
st.title("📦 Sistema de Bipagem e Roteirização - O Boticário")
st.markdown("---")
st.subheader("Bem-vindo!")
st.write("""
Use o menu na barra lateral para navegar entre as páginas:

- **🏠 Home:** Esta página inicial.
- **📱 Scanner:** Use esta página para registrar (bipar) as caixas que chegam.
- **📊 Dashboard:** Use esta página para visualizar os itens bipados, atribuir rotas e acompanhar o progresso.
""")

st.info("O sistema está pronto. Selecione uma página na barra lateral para começar.")
