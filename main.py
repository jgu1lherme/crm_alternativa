import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.express as px

# Configuração da página
st.set_page_config(
    page_title="BI - Alternativa Distribuidora", page_icon="logo2.png", layout="wide"
)

# Barra lateral para navegação
menu = st.sidebar.radio(
    "Selecione uma opção:", ["CRM de Clientes", "Consultor de Meta de CNPJ"]
)

# Aplicação de estilo para fundo da página
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0e1117;
    }
    h1, h2, h3, p, label {
        color: white;
    }
    .stDataFrame {
        background-color: #1c2025;
        border-radius: 10px;
        padding: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Função para carregar arquivos por abas
if "uploaded_file_crm" not in st.session_state:
    st.session_state.uploaded_file_crm = None
if "uploaded_file_cnpj" not in st.session_state:
    st.session_state.uploaded_file_cnpj = None

if menu == "CRM de Clientes":
    st.title("📊 CRM de Clientes - Ativos e Inativos")
    uploaded_file = st.file_uploader(
        "📂 Envie a planilha Excel", type=["xlsx", "xls"], key="crm"
    )
    if uploaded_file:
        st.session_state.uploaded_file_crm = uploaded_file

    df = None
    if st.session_state.uploaded_file_crm:
        df = pd.read_excel(st.session_state.uploaded_file_crm)

    if df is not None:
        df["NFS_EMISSAO"] = pd.to_datetime(df["NFS_EMISSAO"])
        hoje = datetime.today()
        tres_meses_atras = hoje - timedelta(days=90)

        # Filtro de vendedor na página principal
        vendedores = df["VEND_NOME"].unique().tolist()
        vendedor_selecionado = st.selectbox(
            "Selecione um Vendedor", ["Todos"] + vendedores
        )

        if vendedor_selecionado != "Todos":
            df = df[df["VEND_NOME"] == vendedor_selecionado]

        clientes = (
            df.groupby("CLI_RAZ")
            .agg(
                ULTIMA_COMPRA=("NFS_EMISSAO", "max"),
                TOTAL_TRIMESTRAL=(
                    "NFS_CUSTO",
                    lambda x: x[df["NFS_EMISSAO"] >= tres_meses_atras].sum(),
                ),
            )
            .reset_index()
        )

        clientes.rename(columns={"CLI_RAZ": "CLIENTES"}, inplace=True)
        clientes["SITUAÇÃO"] = clientes["ULTIMA_COMPRA"].apply(
            lambda x: "🟢 Ativo" if x >= tres_meses_atras else "🔴 Inativo"
        )

        st.markdown("### 📋 Dados dos Clientes")
        st.dataframe(
            clientes.style.format(
                {
                    "ULTIMA_COMPRA": lambda x: x.strftime("%d/%m/%Y"),
                    "TOTAL_TRIMESTRAL": "R$ {:,.2f}".format,
                }
            ),
            # use_container_width=True,
        )

        ativos = clientes[clientes["SITUAÇÃO"] == "🟢 Ativo"].shape[0]
        inativos = clientes[clientes["SITUAÇÃO"] == "🔴 Inativo"].shape[0]

        fig = px.pie(
            values=[ativos, inativos],
            names=["Ativos", "Inativos"],
            title="Distribuição de Clientes",
        )
        st.plotly_chart(fig)
        # use_container_width=True

        st.success(f"✅ Clientes Ativos: {ativos}")
        st.error(f"❌ Clientes Inativos: {inativos}")
    else:
        st.warning("⚠️ Por favor, envie um arquivo Excel para visualizar os dados.")

elif menu == "Consultor de Meta de CNPJ":
    st.title("📈 Consultor de Meta de CNPJ")
    uploaded_file = st.file_uploader(
        "📂 Envie a planilha Excel contendo CNPJs", type=["xlsx", "xls"], key="cnpj"
    )
    if uploaded_file:
        st.session_state.uploaded_file_cnpj = uploaded_file

    df = None
    if st.session_state.uploaded_file_cnpj:
        df = pd.read_excel(st.session_state.uploaded_file_cnpj)

    if df is not None:
        if "CLI_CGCCPF" not in df.columns:
            st.error("A planilha deve conter a coluna 'CLI_CGCCPF'")
        else:
            data_final = st.date_input(
                "Selecione a data limite para atingir a meta:",
                datetime.today() + timedelta(days=30),
            )

            cnpjs_unicos = df["CLI_CGCCPF"].drop_duplicates()
            total_unicos = len(cnpjs_unicos)
            meta = 600
            dias_uteis_restantes = max((data_final - datetime.today().date()).days, 1)

            restante = meta - total_unicos
            media_diaria = restante / dias_uteis_restantes

            if total_unicos >= meta:
                st.success(f"🎉 Parabéns! Meta atingida ({total_unicos}/{meta})")
            else:
                st.warning(
                    f"📊 Faltam {restante} CNPJs para atingir a meta ({total_unicos}/{meta}).\n\n"
                    f"Você precisa cadastrar {media_diaria:.1f} CNPJs por dia até {data_final.strftime('%d/%m/%Y')}"
                )

            fig = px.bar(
                x=["Meta", "Realizado"],
                y=[meta, total_unicos],
                color=["Meta", "Realizado"],
                color_discrete_sequence=["gray", "#fc630b"],
                title="Progresso da Meta",
            )
            st.plotly_chart(fig)
            # use_container_width=True
    else:
        st.warning("⚠️ Por favor, envie um arquivo Excel para visualizar os dados.")
