import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, date
import io
import os
import json

# Configuração da Página (Responsiva para Desktop e Celular)
st.set_page_config(
    page_title="Gestão do Condomínio | Painel do Síndico",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Diretório para salvar comprovantes (notas e boletos)
UPLOAD_DIR = "comprovantes_digitais"
DB_FILE = "dados_condominio.xlsx"
os.makedirs(UPLOAD_DIR, exist_ok=True)

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

# Funções de Carga e Persistência
def carregar_dados():
    if not os.path.exists(DB_FILE):
        df_lan = pd.DataFrame(columns=[
            "ID", "Data", "Competência", "Tipo", "Conta", "Categoria",
            "Fornecedor / Pagador", "Descrição Detalhada", "Valor (R$)",
            "Status Comprovante", "Nº NF / Doc", "Origem Registro", "Link Comprovante (Drive)"
        ])
        df_man = pd.DataFrame(columns=[
            "ID Manut.", "Ano", "Mês / Competência", "Data Execução",
            "Serviço / Intervenção Realizada", "Área / Equipamento", "ID Despesa Vinculada"
        ])
        return df_lan, df_man

    df_lan = pd.read_excel(DB_FILE, sheet_name="Lancamentos", skiprows=3)
    df_lan = df_lan.dropna(subset=["ID"])
    df_man = pd.read_excel(DB_FILE, sheet_name="Historico_Manutencoes", skiprows=2)
    df_man = df_man.dropna(subset=["ID Manut."])
    return df_lan, df_man

def salvar_dados(df_lan, df_man):
    with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        df_lan.to_excel(writer, sheet_name="Lancamentos", startrow=3, index=False)
        df_man.to_excel(writer, sheet_name="Historico_Manutencoes", startrow=2, index=False)

# Sidebar de Navegação
st.sidebar.title("🏢 Gestão Condominial")
st.sidebar.caption("Acesso Desktop, Mobile & WhatsApp IA")

menu = st.sidebar.radio(
    "Navegação",
    [
        "📊 Dashboard & Classificações",
        "➕ Novo Lançamento (Nota/Boleto)",
        "🔧 Histórico de Manutenções",
        "📑 Relatório Administradora (Excel)",
        "📂 Repositório Digital"
    ]
)

df_lan, df_man = carregar_dados()

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

    # Indicadores Principais
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

    st.subheader("Extrato Detalhado Filtrado")
    st.dataframe(df_filt, use_container_width=True, hide_index=True)

# ==============================================================================
# TELA 2: NOVO LANÇAMENTO (DESKTOP OU CÂMERA DO CELULAR)
# ==============================================================================
elif menu == "➕ Novo Lançamento (Nota/Boleto)":
    st.title("➕ Registrar Movimentação & Comprovante")
    st.caption("No celular, toque em 'Browse files' para tirar foto direta da nota fiscal ou boleto.")

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

        submitted = st.form_submit_button("💾 Salvar Lançamento", use_container_width=True)

        if submitted:
            novo_id = f"LANC-{len(df_lan) + 1:03d}"
            comp_str = dt_lanc.strftime("%Y-%m")
            link_arquivo = "Pendente Upload"

            if arquivo is not None:
                ext = arquivo.name.split(".")[-1]
                nome_limpo = f"{dt_lanc.strftime('%Y-%m-%d')}_{fornecedor.replace(' ', '_')}_{novo_id}.{ext}"
                caminho_salvo = os.path.join(UPLOAD_DIR, nome_limpo)
                with open(caminho_salvo, "wb") as f:
                    f.write(arquivo.getbuffer())
                link_arquivo = caminho_salvo
                status_comp = "OK (Anexado)"

            nova_linha = pd.DataFrame([{
                "ID": novo_id,
                "Data": pd.to_datetime(dt_lanc),
                "Competência": comp_str,
                "Tipo": tipo,
                "Conta": conta,
                "Categoria": categoria,
                "Fornecedor / Pagador": fornecedor,
                "Descrição Detalhada": descricao,
                "Valor (R$)": valor,
                "Status Comprovante": status_comp,
                "Nº NF / Doc": num_nf if num_nf else "-",
                "Origem Registro": "App Web/Mobile",
                "Link Comprovante (Drive)": link_arquivo
            }])

            df_lan = pd.concat([df_lan, nova_linha], ignore_index=True)

            if vincular_manut:
                novo_man_id = f"MAN-{len(df_man) + 1:03d}"
                nova_man = pd.DataFrame([{
                    "ID Manut.": novo_man_id,
                    "Ano": dt_lanc.year,
                    "Mês / Competência": dt_lanc.strftime("%B/%Y"),
                    "Data Execução": pd.to_datetime(dt_lanc),
                    "Serviço / Intervenção Realizada": f"{descricao} ({fornecedor})",
                    "Área / Equipamento": area_manut,
                    "ID Despesa Vinculada": novo_id
                }])
                df_man = pd.concat([df_man, nova_man], ignore_index=True)

            salvar_dados(df_lan, df_man)
            st.success(f"Lançamento {novo_id} registrado com sucesso!")

# ==============================================================================
# TELA 3: HISTÓRICO DE MANUTENÇÕES (2022 - 2026+)
# ==============================================================================
elif menu == "🔧 Histórico de Manutenções":
    st.title("🔧 Histórico de Manutenções & Garantias")

    c1, c2, c3 = st.columns(3)
    anos = ["Todos"] + sorted(df_man["Ano"].dropna().astype(int).unique().tolist(), reverse=True)
    ano_sel = c1.selectbox("Ano", anos)
    area_sel = c2.selectbox("Área / Equipamento", ["Todas"] + AREAS_MANUTENCAO)
    busca = c3.text_input("Buscar serviço ou fornecedor", placeholder="Ex.: bomba, piscina, Otis, 1001...")

    df_m_filt = df_man.copy()
    if ano_sel != "Todos":
        df_m_filt = df_m_filt[df_m_filt["Ano"] == ano_sel]
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
    st.write("Gere e baixe a planilha Excel formatada pronta para enviar à administradora validar.")

    competencias = sorted(df_lan["Competência"].astype(str).unique().tolist(), reverse=True)
    comp_export = st.selectbox("Selecione a Competência para Fechamento", competencias)

    col_s1, col_s2 = st.columns(2)
    saldo_ini_cc = col_s1.number_input("Saldo Inicial Conta Corrente (R$)", value=74688.25, format="%.2f")
    saldo_ini_cx = col_s2.number_input("Saldo Inicial Conta Caixa (R$)", value=1.00, format="%.2f")

    if os.path.exists(DB_FILE):
        with open(DB_FILE, "rb") as f:
            st.download_button(
                label=f"📥 Baixar Planilha Prestação de Contas ({comp_export}).xlsx",
                data=f.read(),
                file_name=f"Prestacao_Contas_Condominio_{comp_export}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ==============================================================================
# TELA 5: REPOSITÓRIO DIGITAL DE COMPROVANTES
# ==============================================================================
elif menu == "📂 Repositório Digital":
    st.title("📂 Repositório de Notas Fiscais e Boletos")
    arquivos = sorted(os.listdir(UPLOAD_DIR), reverse=True)
    if not arquivos:
        st.info("Nenhum comprovante digital enviado ainda pelo app. Os arquivos enviados aparecerão aqui organizados por data.")
    else:
        for arq in arquivos:
            caminho = os.path.join(UPLOAD_DIR, arq)
            with open(caminho, "rb") as f:
                st.download_button(f"📄 {arq}", data=f.read(), file_name=arq)
