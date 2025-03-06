import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import plotly.express as px
import PyPDF2
import streamlit as st

# Configuração da página
st.set_page_config(
    page_title="BI - Alternativa Distribuidora", page_icon="logo2.png", layout="wide"
)

if "uploaded_file_crm" not in st.session_state:
    st.session_state.uploaded_file_crm = None

if "uploaded_file_cnpj" not in st.session_state:
    st.session_state.uploaded_file_cnpj = None

# Barra lateral para navegação
menu = st.sidebar.radio(
    "Selecione uma opção:",
    ["CRM de Clientes", "Positivação de CNPJ", "Renomear Notas Fiscais"],
)

# 🟢 FUNÇÕES DE RENOMEAÇÃO DE NOTAS


# Função para extrair PDFs do ZIP enviado
def extract_pdfs_from_zip(zip_file):
    extracted_pdfs = []
    with zipfile.ZipFile(zip_file, "r") as z:
        for file_name in z.namelist():
            if file_name.lower().endswith(".pdf"):  # Apenas arquivos .pdf
                with z.open(file_name) as f:
                    pdf_bytes = f.read()
                    extracted_pdfs.append((file_name, pdf_bytes))
    return extracted_pdfs


# Função para extrair informações do PDF e gerar nome novo
def extract_info_from_pdf(pdf_bytes):
    try:
        reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
        text = "\n".join(
            page.extract_text() for page in reader.pages if page.extract_text()
        )

        emitente_match = re.search(
            r"IDENTIFICAÇÃO DO EMITENTE\s*([\wÀ-ÿ\-.,& ]+)", text, re.MULTILINE
        )
        numero_match = re.search(r"Nº\.:\s*(\d{3}\.\d{3}\.\d{3})", text)

        if emitente_match and numero_match:
            emitente = emitente_match.group(1).strip()
            numero_nota = numero_match.group(1).strip()
            return f"{numero_nota} - {emitente}.pdf"
    except Exception as e:
        print(f"Erro ao processar PDF: {e}")
    return None


# 🟢 MENU "RENOMEAR NOTAS FISCAIS"
if menu == "Renomear Notas Fiscais":
    st.title("📑 Renomeador de Notas Fiscais")

    # Opção de envio: ZIP ou PDFs individuais
    tipo_upload = st.radio(
        "Escolha como enviar os arquivos:", ["ZIP com PDFs", "Arquivos PDF individuais"]
    )

    pdfs = []

    if tipo_upload == "ZIP com PDFs":
        uploaded_zip = st.file_uploader("📂 Envie um arquivo ZIP", type=["zip"])
        if uploaded_zip:
            with st.spinner("Extraindo arquivos..."):
                pdfs = extract_pdfs_from_zip(uploaded_zip)

    elif tipo_upload == "Arquivos PDF individuais":
        uploaded_pdfs = st.file_uploader(
            "📂 Selecione um ou mais PDFs", type=["pdf"], accept_multiple_files=True
        )
        if uploaded_pdfs:
            pdfs = [(file.name, file.read()) for file in uploaded_pdfs]

    # Processamento dos arquivos enviados
    if pdfs:
        with st.spinner("Processando arquivos..."):
            renamed_data = []  # Lista de PDFs renomeados

            for original_name, pdf_bytes in pdfs:
                new_name = extract_info_from_pdf(pdf_bytes)

                if new_name:
                    renamed_data.append((new_name, pdf_bytes))  # Salvar nome e conteúdo
                else:
                    st.warning(f"⚠️ Não foi possível renomear: {original_name}")

            # Exibir lista de arquivos renomeados
            if renamed_data:
                st.success("✅ PDFs renomeados com sucesso!")
                st.write("### 📋 Arquivos disponíveis para download:")

                for file_name, pdf_bytes in renamed_data:
                    col1, col2 = st.columns([4, 1])
                    col1.write(f"📄 {file_name}")  # Exibir nome do arquivo
                    col2.download_button(
                        label="📥 Baixar",
                        data=pdf_bytes,
                        file_name=file_name,
                        mime="application/pdf",
                    )

                # Criar ZIP para baixar todos os arquivos juntos
                with st.spinner("Criando arquivo ZIP..."):
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as z:
                        for file_name, pdf_bytes in renamed_data:
                            z.writestr(file_name, pdf_bytes)
                    zip_buffer.seek(0)

                st.markdown("### 📂 Baixar todos os arquivos:")
                st.download_button(
                    label="📥 Baixar Tudo (ZIP)",
                    data=zip_buffer,
                    file_name="Notas_Renomeadas.zip",
                    mime="application/zip",
                )

            else:
                st.error("⚠️ Nenhum arquivo foi renomeado.")

# Outros menus existentes
elif menu == "CRM de Clientes":
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
        )

        ativos = clientes[clientes["SITUAÇÃO"] == "🟢 Ativo"].shape[0]
        inativos = clientes[clientes["SITUAÇÃO"] == "🔴 Inativo"].shape[0]

        fig = px.pie(
            values=[ativos, inativos],
            names=["Ativos", "Inativos"],
            title="Distribuição de Clientes",
        )
        st.plotly_chart(fig)

        st.success(f"✅ Clientes Ativos: {ativos}")
        st.error(f"❌ Clientes Inativos: {inativos}")
    else:
        st.warning("⚠️ Por favor, envie um arquivo Excel para visualizar os dados.")

elif menu == "Positivação de CNPJ":
    st.title("📈 Positivação de CNPJ")
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
    else:
        st.warning("⚠️ Por favor, envie um arquivo Excel para visualizar os dados.")
