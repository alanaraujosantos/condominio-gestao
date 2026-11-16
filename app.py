import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from datetime import datetime, date
import io
import openpyxl

# Configuração da Página (Responsiva para Desktop e Celular)
st.set_page_config(
    page_title="Gestão do Condomínio | Painel do Síndico",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

CATEGORIAS = [
    "Recebimento de Taxas Condominiais",
    "Concessionárias (Água, Gás, Energia, Internet)",
    "Contratos Fixos & Terceirizados",
    "Manutenção Corretiva & Conservação",
    "Piscina & Jardinagem",
    "Investimentos / Benfeitorias",
    "Despesas Bancárias & Reembolsos",
    "Aplicação Financeira (RDC)"
]

CONTAS = ["Conta Corrente", "Conta Caixa", "RDC Fundo Reserva", "RDC Fundo Fachada"]
AREAS_MANUTENCAO = [
    "Piscina & Lazer", "Portões & Acessos", "Elevadores",
    "Hidráulica & Bombas", "Elétrica & Iluminação",
    "Estrutura & Fachada", "Segurança & Telecom",
    "Conservação & Segurança", "Áreas Comuns"
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Conexão com Google Sheets e Google Drive via Secrets do Streamlit
@st.cache_resource
def conectar_google():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    drive_service = build("drive", "v3", credentials=creds)
    return gc, drive_service

def carregar_dados():
    gc, _ = conectar_google()
    sh = gc.open_by_key(st.secrets["PLANILHA_ID"])

    ws_lan = sh.worksheet("Lancamentos")
    vals_lan = ws_lan.get_all_values()
    if len(vals_lan) > 4:
        headers_lan = vals_lan[3]
        df_lan = pd.DataFrame(vals_lan[4:], columns=headers_lan)
        df_lan = df_lan[df_lan["ID"].astype(str).str.strip() != ""]
        df_lan["Valor (R$)"] = (
            df_lan["Valor (R$)"]
            .astype(str)
            .str.replace("R$", "", regex=False)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .str.strip()
        )
        df_lan["Valor (R$)"] = pd.to_numeric(df_lan["Valor (R$)"], errors="coerce").fillna(0.0)
    else:
        df_lan = pd.DataFrame(columns=[
            "ID", "Data", "Competência", "Tipo", "Conta", "Categoria",
            "Fornecedor / Pagador", "Descrição Detalhada", "Valor (R$)",
            "Status Comprovante", "Nº NF / Doc", "Origem Registro", "Link Comprovante (Drive)"
        ])

    ws_man = sh.worksheet("Historico_Manutencoes")
    vals_man = ws_man.get_all_values()
    if len(vals_man) > 3:
        headers_man = vals_man[2]
        df_man = pd.DataFrame(vals_man[3:], columns=headers_man)
        df_man = df_man[df_man["ID Manut."].astype(str).str.strip() != ""]
    else:
        df_man = pd.DataFrame(columns=[
            "ID Manut.", "Ano", "Mês / Competência", "Data Execução",
            "Serviço / Intervenção Realizada", "Área / Equipamento", "ID Despesa Vinculada"
        ])

    return df_lan, df_man

def upload_comprovante_drive(arquivo, nome_arquivo):
    _, drive_service = conectar_google()
    pasta_id = st.secrets["PASTA_DRIVE_ID"]
    file_metadata = {"name": nome_arquivo, "parents": [pasta_id]}
    media = MediaIoBaseUpload(io.BytesIO(arquivo.getbuffer()), mimetype=arquivo.type, resumable=True)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()
    return file.get("webViewLink")

def adicionar_lancamento_sheets(nova_linha_lista):
    gc, _ = conectar_google()
    sh = gc.open_by_key(st.secrets["PLANILHA_ID"])
    ws_lan = sh.worksheet("Lancamentos")
    ws_lan.append_row(nova_linha_lista, value_input_option="USER_ENTERED")

def adicionar_manutencao_sheets(nova_man_lista):
    gc, _ = conectar_google()
    sh = gc.open_by_key(st.secrets["PLANILHA_ID"])
    ws_man = sh.worksheet("Historico_Manutencoes")
    ws_man.append_row(nova_man_lista, value_input_option="USER_ENTERED")

# Sidebar de Navegação
st.sidebar.title("🏢 Gestão Condominial")
st.sidebar.caption("Integrado ao Google Sheets & Drive")

if st.sidebar.button("🔄 Atualizar Dados"):
    st.rerun()

menu = st.sidebar.radio(
    "Navegação",
    [
        "📊 Dashboard & Classificações",
        "➕ Novo Lançamento (Nota/Boleto)",
        "🔧 Histórico de Manutenções",
        "📑 Relatório Administradora (Excel)"
    ]
)

try:
    df_lan, df_man = carregar_dados()
except Exception as e:
    st.error(f"Erro ao conectar à planilha Google: {e}")
    st.info("Verifique se os Secrets (gcp_service_account, PLANILHA_ID e PASTA_DRIVE_ID) foram configurados no Streamlit.")
    st.stop()

# ==============================================================================
# TELA 1: DASHBOARD & CLASSIFICAÇÕES
# ==============================================================================
if menu == "📊 Dashboard & Classificações":
    st.title("📊 Painel Analítico do Condomínio")

    competencias = sorted(df_lan["Competência"].astype(str).unique().tolist(), reverse=True)
    col_f1, col_f2, col_f3 = st.columns(3)
    comp_sel = col_f1.selectbox("Competência (Mês/Ano)", ["Todas"] + competencias)
    conta_sel = col_f2.selectbox("Conta", ["Todas"] + CONTAS)
    cat_sel = col_f3.selectbox("Categoria", ["Todas"] + CATEGORIAS)

    df_filt = df_lan.copy()
    if comp_sel != "Todas":
        df_filt = df_filt[df_filt["Competência"].astype(str) == comp_sel]
    if conta_sel != "Todas":
        df_filt = df_filt[df_filt["Conta"] == conta_sel]
    if cat_sel != "Todas":
        df_filt = df_filt[df_filt["Categoria"] == cat_sel]

    receitas = df_filt[df_filt["Tipo"] == "Receita"]["Valor (R$)"].sum()
    despesas_tot = df_filt[df_filt["Tipo"] == "Despesa"]["Valor (R$)"].sum()
    investimentos = df_filt[
        (df_filt["Tipo"] == "Despesa") & (df_filt["Categoria"] == "Investimentos / Benfeitorias")
    ]["Valor (R$)"].sum()
    custeio_ord = despesas_tot - investimentos
    pendencias = len(df_filt[df_filt["Status Comprovante"] == "FALTA RECIBO"])

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Receitas (Taxas)", f"R$ {receitas:,.2f}")
    k2.metric("Despesas Ordinárias", f"R$ {custeio_ord:,.2f}")
    k3.metric("Investimentos / Melhorias", f"R$ {investimentos:,.2f}")
    k4.metric("Pendências de Recibo", f"{pendencias} item(ns)", delta_color="inverse")

    st.divider()
    c_left, c_right = st.columns([1.2, 1])

    with c_left:
        st.subheader("Gastos por Categoria")
        df_desp = df_filt[df_filt["Tipo"] == "Despesa"]
        if not df_desp.empty:
            grp = df_desp.groupby("Categoria")["Valor (R$)"].sum().sort_values(ascending=False)
            st.bar_chart(grp)
        else:
            st.info("Nenhuma despesa encontrada para os filtros selecionados.")

    with c_right:
        st.subheader("⚠️ Alertas de Auditoria (Falta Recibo)")
        df_pend = df_filt[df_filt["Status Comprovante"] == "FALTA RECIBO"]
        if df_pend.empty:
            st.success("Todos os lançamentos filtrados possuem comprovante anexado!")
        else:
            st.dataframe(
                df_pend[["Data", "Fornecedor / Pagador", "Descrição Detalhada", "Valor (R$)"]],
                use_container_width=True,
                hide_index=True
            )

    st.subheader("Extrato Detalhado (com Links dos Comprovantes)")
    st.dataframe(
        df_filt,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Link Comprovante (Drive)": st.column_config.LinkColumn("Comprovante")
        }
    )

# ==============================================================================
# TELA 2: NOVO LANÇAMENTO (DESKTOP OU CÂMERA DO CELULAR -> GOOGLE DRIVE)
# ==============================================================================
elif menu == "➕ Novo Lançamento (Nota/Boleto)":
    st.title("➕ Registrar Movimentação & Comprovante")
    st.caption("Salva direto na Planilha Google e envia o comprovante para a pasta do Google Drive.")

    with st.form("form_lancamento", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        dt_lanc = c1.date_input("Data", value=date.today())
        tipo = c2.selectbox("Tipo de Movimentação", ["Despesa", "Receita", "Transferência / Aplicação"])
        conta = c3.selectbox("Conta", CONTAS)

        c4, c5 = st.columns(2)
        categoria = c4.selectbox("Categoria Contábil", CATEGORIAS)
        fornecedor = c5.text_input("Fornecedor / Prestador / Pagador", placeholder="Ex.: Otis, Emerson, Cagepa...")

        descricao = st.text_input("Descrição Detalhada", placeholder="Ex.: Substituição base antena e reparo portão")

        c6, c7, c8 = st.columns(3)
        valor = c6.number_input("Valor (R$)", min_value=0.0, step=10.0, format="%.2f")
        num_nf = c7.text_input("Nº Nota Fiscal / Boleto", placeholder="Ex.: NF 44")
        status_comp = c8.selectbox("Status do Comprovante", ["OK (Anexado)", "FALTA RECIBO", "OK (Extrato)"])

        arquivo = st.file_uploader("Anexar Nota Fiscal / Boleto (PDF, JPG, PNG)", type=["pdf", "jpg", "jpeg", "png"])

        vincular_manut = st.checkbox("Registrar também no Histórico de Manutenções?")
        area_manut = st.selectbox("Área / Equipamento (se for manutenção)", AREAS_MANUTENCAO)

        submitted = st.form_submit_button("💾 Salvar na Nuvem", use_container_width=True)

        if submitted:
            with st.spinner("Salvando lançamento e enviando arquivo ao Google Drive..."):
                novo_id = f"LANC-{len(df_lan) + 1:03d}"
                comp_str = dt_lanc.strftime("%Y-%m")
                link_arquivo = "Pendente Upload"

                if arquivo is not None:
                    ext = arquivo.name.split(".")[-1]
                    forn_limpo = fornecedor.strip().replace(" ", "_") if fornecedor else "SemFornecedor"
                    nome_limpo = f"{dt_lanc.strftime('%Y-%m-%d')}_{forn_limpo}_{novo_id}.{ext}"
                    link_arquivo = upload_comprovante_drive(arquivo, nome_limpo)
                    status_comp = "OK (Anexado)"

                nova_linha = [
                    novo_id,
                    dt_lanc.strftime("%d/%m/%Y"),
                    comp_str,
                    tipo,
                    conta,
                    categoria,
                    fornecedor,
                    descricao,
                    valor,
                    status_comp,
                    num_nf if num_nf else "-",
                    "App Web/Mobile",
                    link_arquivo
                ]
                adicionar_lancamento_sheets(nova_linha)

                if vincular_manut:
                    novo_man_id = f"MAN-{len(df_man) + 1:03d}"
                    nova_man = [
                        novo_man_id,
                        dt_lanc.year,
                        dt_lanc.strftime("%B/%Y"),
                        dt_lanc.strftime("%d/%m/%Y"),
                        f"{descricao} ({fornecedor})",
                        area_manut,
                        novo_id
                    ]
                    adicionar_manutencao_sheets(nova_man)

                st.success(f"Lançamento {novo_id} gravado com sucesso no Google Sheets!")

# ==============================================================================
# TELA 3: HISTÓRICO DE MANUTENÇÕES (2022 - 2026+)
# ==============================================================================
elif menu == "🔧 Histórico de Manutenções":
    st.title("🔧 Histórico de Manutenções & Garantias")

    c1, c2, c3 = st.columns(3)
    anos = ["Todos"] + sorted(df_man["Ano"].astype(str).unique().tolist(), reverse=True)
    ano_sel = c1.selectbox("Ano", anos)
    area_sel = c2.selectbox("Área / Equipamento", ["Todas"] + AREAS_MANUTENCAO)
    busca = c3.text_input("Buscar serviço ou fornecedor", placeholder="Ex.: bomba, piscina, Otis, 1001...")

    df_m_filt = df_man.copy()
    if ano_sel != "Todos":
        df_m_filt = df_m_filt[df_m_filt["Ano"].astype(str) == ano_sel]
    if area_sel != "Todas":
        df_m_filt = df_m_filt[df_m_filt["Área / Equipamento"] == area_sel]
    if busca:
        df_m_filt = df_m_filt[
            df_m_filt["Serviço / Intervenção Realizada"].astype(str).str.contains(busca, case=False, na=False)
        ]

    st.metric("Total de Intervenções Encontradas", len(df_m_filt))
    st.dataframe(df_m_filt, use_container_width=True, hide_index=True)

# ==============================================================================
# TELA 4: RELATÓRIO ADMINISTRADORA (EXPORTAÇÃO EXCEL)
# ==============================================================================
elif menu == "📑 Relatório Administradora (Excel)":
    st.title("📑 Fechamento Mensal para Administradora")
    st.write("Gere o arquivo Excel atualizado diretamente da base na nuvem para enviar à administradora.")

    competencias = sorted(df_lan["Competência"].astype(str).unique().tolist(), reverse=True)
    comp_export = st.selectbox("Selecione a Competência para Fechamento", competencias)

    df_mes = df_lan[df_lan["Competência"].astype(str) == comp_export].copy()
    st.subheader(f"Prévia dos Lançamentos — {comp_export}")
    st.dataframe(df_mes, use_container_width=True, hide_index=True)

    # Gera arquivo Excel em memória com as abas atualizadas
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_mes.to_excel(writer, sheet_name=f"Prestacao_{comp_export}", index=False)
        df_lan.to_excel(writer, sheet_name="Base_Completa_Lancamentos", index=False)
        df_man.to_excel(writer, sheet_name="Historico_Manutencoes", index=False)

    st.download_button(
        label=f"📥 Baixar Planilha Prestação de Contas ({comp_export}).xlsx",
        data=output.getvalue(),
        file_name=f"Prestacao_Contas_Condominio_{comp_export}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
