import json
import os
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from google_utils import get_gspread_client

load_dotenv()

st.set_page_config(layout="wide")
st.title("📬 Gerador de mensagens de acareação")

STATE_FILE = Path(__file__).with_name("acareacoes_sent_state.json")
NOME_PLANILHA = os.getenv("NOME_PLANILHA", "acareaBase")


def carregar_estado() -> dict:
    if not STATE_FILE.exists():
        return {}
    try:
        with STATE_FILE.open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except Exception:
        return {}


def salvar_estado(dados: dict) -> None:
    with STATE_FILE.open("w", encoding="utf-8") as handle:
        json.dump(dados, handle, ensure_ascii=False, indent=2)


def normalizar_enviado(valor) -> bool:
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, (int, float)):
        return bool(valor)
    if valor is None:
        return False
    return str(valor).strip().lower() in {"true", "1", "yes", "y", "sim", "t"}


def atualizar_estado_na_planilha(aba: str, df: pd.DataFrame) -> None:
    if df.empty:
        return

    try:
        cliente = get_gspread_client()
        planilha = cliente.open(NOME_PLANILHA).worksheet(aba)
    except Exception as exc:
        raise RuntimeError(f"Não foi possível abrir a aba {aba}: {exc}") from exc

    rows = planilha.get_all_values()
    if not rows:
        return

    headers = rows[0]
    headers_lower = [str(header).strip().lower() for header in headers]

    def encontrar_indice(*candidatos):
        for nome in candidatos:
            if nome in headers_lower:
                return headers_lower.index(nome)
        return None

    awb_index = encontrar_indice("awb", "codigo", "waybill", "tracking")
    if awb_index is None:
        awb_index = 0

    if "enviado" not in headers_lower:
        planilha.update_cell(1, len(headers) + 1, "Enviado")
        headers.append("Enviado")
        headers_lower = [str(header).strip().lower() for header in headers]

    enviado_index = headers_lower.index("enviado")

    estados = {}
    for _, row in df.iterrows():
        awb = str(row.get("AWB", row.get("codigo", row.get("Waybill", "")))).strip()
        if awb:
            estados[awb] = bool(row.get("Enviado", False))

    for row_number, row in enumerate(rows[1:], start=2):
        if awb_index is not None and awb_index < len(row):
            awb = str(row[awb_index]).strip()
            if awb in estados:
                planilha.update_cell(row_number, enviado_index + 1, "TRUE" if estados[awb] else "FALSE")


def encontrar_planilhas_google() -> list[str]:
    try:
        cliente = get_gspread_client()
        planilha = cliente.open(NOME_PLANILHA)
        return [aba.title for aba in planilha.worksheets]
    except Exception as exc:
        st.warning(f"Não foi possível ler as abas do Google Sheets: {exc}")
        return []


def carregar_dados(aba: str) -> pd.DataFrame:
    try:
        cliente = get_gspread_client()
        planilha = cliente.open(NOME_PLANILHA)
        worksheet = planilha.worksheet(aba)
        records = worksheet.get_all_records()
        if not records:
            return pd.DataFrame()
        df = pd.DataFrame(records)
        df = df.fillna("")
        return df
    except Exception as exc:
        st.error(f"Erro ao carregar a aba {aba}: {exc}")
        return pd.DataFrame()


def preparar_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    colunas = [
        "AWB",
        "Motorista",
        "Nome",
        "Número do Cliente",
        "Telefone",
        "Bairro",
        "Endereco",
        "Produto",
        "Valor",
        "Prazo do Processo",
    ]

    for coluna in colunas:
        if coluna not in df.columns:
            df[coluna] = ""

    df = df[colunas].copy()
    df["AWB"] = df["AWB"].astype(str).str.strip()
    df["Bairro"] = df["Bairro"].astype(str).fillna("")
    df["Número do Cliente"] = df["Número do Cliente"].astype(str).fillna("")
    df["Endereco"] = df["Endereco"].astype(str).fillna("")
    return df


st.subheader("1. Escolha a base")
abas = encontrar_planilhas_google()
if abas:
    aba_selecionada = st.selectbox("Aba do Google Sheets", abas)
else:
    aba_selecionada = st.text_input("Nome da aba", "ITR")

if st.button("Carregar dados") or abas:
    df = carregar_dados(aba_selecionada)
    if df.empty:
        st.info("Nenhuma linha encontrada para essa aba.")
        st.stop()

    df = preparar_df(df)
    estado = carregar_estado()

    if "Enviado" in df.columns:
        df["Enviado"] = df["Enviado"].apply(normalizar_enviado)
    else:
        df["Enviado"] = False

    for index, row in df.iterrows():
        awb = str(row.get("AWB", "")).strip()
        if awb and awb in estado:
            df.at[index, "Enviado"] = bool(estado[awb])

    total = len(df)
    enviados = int(df["Enviado"].sum())
    faltam = total - enviados

    st.metric("Total de acareações", total)
    st.metric("Já enviadas", enviados)
    st.metric("Faltam", faltam)

    st.subheader("2. Gerador de mensagens")
    colunas_visiveis = [
        "Enviado",
        "AWB",
        "Motorista",
        "Nome",
        "Número do Cliente",
        "Bairro",
        "Endereco",
        "Produto",
        "Prazo do Processo",
    ]
    editor = st.data_editor(
        df[colunas_visiveis].copy(),
        hide_index=True,
        use_container_width=True,
        column_config={
            "Enviado": st.column_config.CheckboxColumn("Enviado"),
            "AWB": st.column_config.TextColumn("AWB"),
            "Motorista": st.column_config.TextColumn("Motorista"),
            "Nome": st.column_config.TextColumn("Nome"),
            "Número do Cliente": st.column_config.TextColumn("Número do Cliente"),
            "Bairro": st.column_config.TextColumn("Bairro"),
            "Endereco": st.column_config.TextColumn("Endereco"),
            "Produto": st.column_config.TextColumn("Produto"),
            "Prazo do Processo": st.column_config.TextColumn("Prazo do Processo"),
        },
        key="editor_acareacoes",
    )

    if st.button("Salvar marcações"):
        novo_estado = {}
        for _, row in editor.iterrows():
            awb = str(row.get("AWB", "")).strip()
            if awb:
                novo_estado[awb] = bool(row.get("Enviado", False))
        salvar_estado(novo_estado)
        try:
            atualizar_estado_na_planilha(aba_selecionada, editor.assign(Enviado=editor["Enviado"].apply(bool)))
            st.success("Marcação salva com sucesso na planilha e no estado local.")
        except Exception as exc:
            st.error(f"Não foi possível salvar na planilha: {exc}")

    st.subheader("3. Preview da mensagem")
    linhas_pendentes = editor[editor["Enviado"] == False]
    if linhas_pendentes.empty:
        st.info("Nenhuma acareação pendente.")
    else:
        amostra = linhas_pendentes.head(3).copy()
        for _, row in amostra.iterrows():
            bairro = row.get("Bairro") or "N/A"
            numero = row.get("Número do Cliente") or row.get("Telefone") or "N/A"
            endereco = row.get("Endereco") or "N/A"
            mensagem = (
                f"Olá! Temos uma acareação para o cliente {row.get('Nome', 'N/A')}.\n"
                f"AWB: {row.get('AWB', 'N/A')}\n"
                f"Número do Cliente: {numero}\n"
                f"Endereço: {endereco}\n"
                f"Bairro: {bairro}\n"
                f"Produto: {row.get('Produto', 'N/A')}\n"
                f"Prazo do Processo: {row.get('Prazo do Processo', 'N/A')}"
            )
            st.text_area(f"Mensagem para {row.get('AWB', 'N/A')}", mensagem, height=180)
