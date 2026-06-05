import os
import re
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

COL_MOTORISTA = "Delivery Associate"
COL_STATUS = "Last Scan Type"
COL_AWB = "Waybill No"


def baixar_inventario(page, sigla):
    print(f"[{sigla}] Navegando para Monitoramento -> Monitor do inventário...")
    page.wait_for_timeout(15000)
    page.get_by_text("Monitoramento").first.click(force=True)
    page.wait_for_timeout(3000)
    page.get_by_text("Monitor de operação").first.click(force=True)
    page.wait_for_timeout(3000)
    page.get_by_text("Monitor do inventário").first.click(force=True)
    page.wait_for_timeout(8000)

    print(f"[{sigla}] Solicitando extração do inventário...")
    page.locator('.ImileActionButton-root:has-text("extrair")').first.click(force=True)
    page.wait_for_timeout(1000)
    page.locator('li.ImileMenuItem-root:has-text("Exportar tudo")').first.click(force=True)
    page.wait_for_timeout(1000)
    page.locator('.export-button').first.click(force=True)

    print(f"[{sigla}] Aguardando a iMile gerar o arquivo...")
    page.wait_for_timeout(10000)

    print(f"[{sigla}] Baixando o arquivo bruto do inventário...")
    icone_download = page.locator('span.Imile-ButtonIcon-root svg path[d*="M8 0C3.57"]').locator('..').first
    icone_download.click(force=True)

    btn_baixar = page.locator('button:has-text("Baixar")').last
    btn_baixar.wait_for(state="visible", timeout=15000)

    with page.expect_download(timeout=120000) as download_info:
        btn_baixar.click(force=True)

    download = download_info.value
    arquivo = f"Inventario_Vencimento_{sigla}.xlsx"
    download.save_as(arquivo)

    print(f"[{sigla}] ✅ Planilha baixada com sucesso: {arquivo}")
    return arquivo


def gerar_relatorio_em_rota(caminho_arquivo, sigla):
    print("\n" + "="*50)
    print(f"🚚 GERANDO RELATÓRIO 'EM ROTA DE ENTREGA' PARA {sigla}...")
    print("="*50)

    try:
        df = pd.read_excel(caminho_arquivo)

        # Busca dinâmica das colunas
        col_status = next((col for col in df.columns if col.lower() == 'last scan type'), 'Last Scan Type')
        col_motorista = next((col for col in df.columns if col.lower() == 'delivery associate'), 'Delivery Associate')

        # 1. Filtrar apenas os que estão "Em rota de entrega" (ignorando maiúsculas/minúsculas)
        df_rota = df[df[col_status].astype(str).str.contains('Em rota de entrega', case=False, na=False)].copy()

        # 2. Definir as abas (tags) com base na sigla da conta
        if sigla.upper() == 'JML':
            tags = ['INT', 'LOC']
        elif sigla.upper() == 'ITR':
            tags = ['STB', 'ITR']
        else:
            tags = ['INT', 'LOC'] # Padrão de segurança

        nome_arquivo_final = f"Relatorio_Em_Rota_{sigla}.xlsx"
        writer = pd.ExcelWriter(nome_arquivo_final, engine='xlsxwriter')
        workbook = writer.book

        # ==========================================
        # ESTILOS (Mesmo padrão de cores da Tabela Anterior)
        # ==========================================
        fmt_cabecalho = workbook.add_format({
            'bg_color': '#000000', 'font_color': '#FFFFFF', 
            'bold': True, 'align': 'center', 'valign': 'vcenter'
        })
        fmt_dados_texto = workbook.add_format({
            'bg_color': '#5B9BD5', 'font_color': '#FFFFFF', 
            'align': 'left', 'valign': 'vcenter'
        })
        fmt_dados_numero = workbook.add_format({
            'bg_color': '#5B9BD5', 'font_color': '#FFFFFF', 
            'align': 'center', 'valign': 'vcenter'
        })

        # 3. Processar cada tag e criar sua aba
        for tag in tags:
            # Filtra os motoristas que contêm a tag no nome (ex: "INT")
            df_tag = df_rota[df_rota[col_motorista].astype(str).str.contains(tag, na=False, case=False)]
            
            if df_tag.empty:
                print(f"⚠️ Nenhum pacote 'Em rota de entrega' encontrado para a tag {tag} na base {sigla}.")
                continue

            # Conta a quantidade por motorista e reseta o index para virar um DataFrame
            contagem = df_tag[col_motorista].value_counts().reset_index()
            contagem.columns = ['Motorista', 'Quantidade']
            
            # Garante a ordem decrescente (do maior para o menor)
            contagem = contagem.sort_values(by='Quantidade', ascending=False)

            # Cria a aba com o nome da Tag
            worksheet = workbook.add_worksheet(f"{sigla} - {tag}")

            # Escreve os cabeçalhos
            worksheet.write(0, 0, 'Motorista', fmt_cabecalho)
            worksheet.write(0, 1, 'Volumes em Rota', fmt_cabecalho)

            # Preenche os dados linha a linha
            for row_idx, row in contagem.iterrows():
                worksheet.write(row_idx + 1, 0, str(row['Motorista']), fmt_dados_texto)
                worksheet.write(row_idx + 1, 1, int(row['Quantidade']), fmt_dados_numero)

            # Ajusta a largura das colunas
            worksheet.set_column(0, 0, 50) # Coluna de Motoristas bem larga
            worksheet.set_column(1, 1, 18) # Coluna de Quantidade

        writer.close()
        print(f"✅ Relatório de motoristas gerado com sucesso para {sigla}: '{nome_arquivo_final}'")

    except Exception as e:
        print(f"❌ Erro ao formatar relatório Em Rota para {sigla}: {e}")
def processar_conta(sigla, usuario, senha, p):
    print(f"\n{'='*50}\n🚀 INICIANDO CONTA: {sigla}\n{'='*50}")
    if not usuario or not senha:
        print(f"⚠️ Credenciais de {sigla} ausentes. Pulando...")
        return

    browser = p.chromium.launch(headless=False, args=["--start-maximized"])
    context = browser.new_context(no_viewport=True, accept_downloads=True)
    page = context.new_page()

    try:
        print(f"[{sigla}] Fazendo login...")
        page.goto("https://ds-login.imile.com/", wait_until="domcontentloaded")
        page.fill('input[type="text"]', usuario)
        page.fill('input[type="password"]', senha)
        page.get_by_role("button", name=re.compile(r"登录|Conecte-se|Login", re.IGNORECASE)).click()
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(5000)

        try:
            page.locator('.close-icon').first.click(force=True)
        except:
            pass

        arquivo_bruto = baixar_inventario(page, sigla)
        if arquivo_bruto and os.path.exists(arquivo_bruto):
            gerar_relatorio_pandas(arquivo_bruto, sigla)

    except Exception as e:
        print(f"❌ Erro crítico na base {sigla}: {e}")

    finally:
        browser.close()


def main():
    contas = [
        ("JML", os.getenv("IMILE_JML_USER"), os.getenv("IMILE_JML_PASS")),
        ("ITR", os.getenv("IMILE_ITR_USER"), os.getenv("IMILE_ITR_PASS"))
    ]

    with sync_playwright() as p:
        for sigla, usuario, senha in contas:
            processar_conta(sigla, usuario, senha, p)


if __name__ == "__main__":
    main()
