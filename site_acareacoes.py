import streamlit as st
import os

st.set_page_config(layout="wide")
st.title("🕵️‍♀️ Página de Depuração de Ambiente")

st.header("1. Onde os arquivos estão?")
try:
    # O diretório raiz do seu repositório no servidor do Streamlit
    root_dir = "/mount/src/"
    
    st.write(f"Listando arquivos em `{root_dir}`:")
    
    # Lista os arquivos e diretórios na raiz
    files_in_root = os.listdir(root_dir)
    st.code('
'.join(files_in_root))
    
except Exception as e:
    st.error(f"Erro ao tentar listar arquivos: {e}")

st.header("2. O 'requirements.txt' está sendo lido corretamente?")
try:
    # O caminho completo esperado para o arquivo
    req_path = "/mount/src/requirements.txt"
    
    st.write(f"Tentando ler o arquivo em `{req_path}`...")
    
    if os.path.exists(req_path):
        with open(req_path, 'r') as f:
            content = f.read()
        st.success("✅ Arquivo 'requirements.txt' encontrado! Conteúdo:")
        st.code(content)
    else:
        st.error("❌ ERRO CRÍTICO: O arquivo 'requirements.txt' NÃO foi encontrado no diretório raiz do repositório.")
        st.info("Verifique se o arquivo 'requirements.txt' está na pasta principal do seu projeto no GitHub, e não dentro de uma subpasta.")

except Exception as e:
    st.error(f"Erro ao tentar ler o requirements.txt: {e}")
