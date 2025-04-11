import io
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timedelta
from io import BytesIO

import fitz  # PyMuPDF
import pandas as pd
import pdf2image
import plotly.express as px
import PyPDF2
import streamlit as st
from docx import Document
from docx2pdf import convert
from PIL import Image
from PyPDF2 import PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Configuração da página
st.set_page_config(
    page_title="BI - Alternativa Distribuidora", page_icon="logo2.png", layout="wide"
)

if "uploaded_file_crm" not in st.session_state:
    st.session_state.uploaded_file_crm = None

if "uploaded_file_cnpj" not in st.session_state:
    st.session_state.uploaded_file_cnpj = None

if "uploaded_file_bancaria" not in st.session_state:
    st.session_state.uploaded_file_bancaria = None

# Barra lateral para navegação
menu = st.sidebar.radio(
    "Selecione uma opção:",
    [
        "CRM de Clientes",
        "Positivação de CNPJ",
        "Renomear Notas Fiscais",
        "Conversor de Arquivos",
        "Organização Planilha Bancária",
        "Contabilidade - Extrato ML",
    ],
)


# Função para organizar planilha bancária
def process_bank_statement(file):
    # Ler a planilha original
    if file.name.endswith(".xls"):
        df = pd.read_excel(file, dtype=str, engine="xlrd")
    else:
        df = pd.read_excel(file, dtype=str, engine="openpyxl")

    # Remover espaços extras e converter nomes das colunas
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    df.columns = ["Data", "Documento", "Historico", "Valor"]

    # Filtrar linhas vazias ou irrelevantes
    df = df.dropna(subset=["Historico", "Valor"], how="all")
    df = df[~df["Historico"].str.contains("SALDO|====>", na=False, case=False)]

    # Criar colunas de Crédito e Débito
    df["Valor Crédito"] = df["Valor"].str.extract(r"([\d,.]+)C$")[0]
    df["Valor Débito"] = df["Valor"].str.extract(r"([\d,.]+)D$")[0]

    # Remover a coluna original "Valor"
    df = df.drop(columns=["Valor"])

    # Converter para número
    def to_numeric(value):
        if pd.notna(value):
            return float(value.replace(".", "").replace(",", "."))
        return 0.0

    df["Valor Crédito"] = df["Valor Crédito"].apply(to_numeric)
    df["Valor Débito"] = df["Valor Débito"].apply(to_numeric)

    # Calcular totais
    total_credito = df["Valor Crédito"].sum()
    total_debito = df["Valor Débito"].sum()
    diferenca = total_credito - total_debito

    # Aplicar formato contábil
    df["Valor Crédito"] = df["Valor Crédito"].map(
        lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )
    df["Valor Débito"] = df["Valor Débito"].map(
        lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )

    # Criar um DataFrame com os totais e a diferença
    total_df = pd.DataFrame(
        {
            "Data": [""],
            "Documento": [""],
            "Historico": ["TOTAL"],
            "Valor Crédito": [
                f"R$ {total_credito:,.2f}".replace(",", "X")
                .replace(".", ",")
                .replace("X", ".")
            ],
            "Valor Débito": [
                f"R$ {total_debito:,.2f}".replace(",", "X")
                .replace(".", ",")
                .replace("X", ".")
            ],
        }
    )

    diferenca_df = pd.DataFrame(
        {
            "Data": [""],
            "Documento": [""],
            "Historico": [
                f"DIFERENÇA (Crédito - Débito): R$ {diferenca:,.2f}".replace(",", "X")
                .replace(".", ",")
                .replace("X", ".")
            ],
            "Valor Crédito": [""],
            "Valor Débito": [""],
        }
    )

    # Concatenar os totais ao final do DataFrame
    df = pd.concat([df, total_df, diferenca_df], ignore_index=True)

    # Salvar a planilha processada em um buffer
    output = io.BytesIO()
    df.to_excel(output, sheet_name="Dados Processados", index=False)
    output.seek(0)

    return output, df
    

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
                extracted_info = extract_info_from_pdf(pdf_bytes)

                if extracted_info:
                    # Inverter a ordem para "Nome - Número"
                    numero, nome = extracted_info.split(" - ", 1)
                    new_name = f"{nome} - {numero}.pdf"
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

# 🟢 MENU "CONVERSOR DE IMAGENS"
elif menu == "Conversor de Arquivos":
    st.title("🖼️ Conversor de Arquivos")

    # Opção de envio de arquivo
    uploaded_file = st.file_uploader(
        "📂 Selecione um arquivo para conversão", type=["png", "jpg", "jpeg", "pdf"]
    )

    # Verificar se o usuário enviou um arquivo
    if uploaded_file:
        # Identificar o tipo de arquivo
        file_extension = uploaded_file.name.split(".")[-1].lower()

        # 🟢 CONVERSÃO PARA IMAGENS (SE FOR UM PDF)
        if file_extension == "pdf":
            st.subheader("Conversão de PDF para Imagens")

            if st.button("Converter PDF para Imagens"):
                try:
                    # Converter o PDF para imagens
                    images = pdf2image.convert_from_bytes(uploaded_file.read())

                    st.success("✅ PDF convertido para imagens com sucesso!")

                    # Disponibilizar cada página do PDF como imagem para download
                    for i, image in enumerate(images):
                        image_io = io.BytesIO()
                        image.save(image_io, "PNG")
                        image_io.seek(0)
                        st.download_button(
                            label=f"📥 Baixar Página {i + 1} (Imagem)",
                            data=image_io,
                            file_name=f"pagina_{i + 1}.png",
                            mime="image/png",
                        )
                except Exception as e:
                    st.error(f"⚠️ Erro ao converter PDF para imagens: {e}")

        # 🟢 CONVERSÃO DE IMAGEM PARA VÁRIOS FORMATOS E PDF (SE FOR UMA IMAGEM)
        elif file_extension in ["png", "jpg", "jpeg"]:
            st.subheader("Conversão de Imagem")

            # Seleção de formatos de conversão, incluindo "JPEG" e "JPG"
            formato_destino = st.selectbox(
                "Escolha o formato para conversão:", ["JPEG", "JPG", "PNG", "PDF"]
            )

            if st.button("Converter Imagem"):
                try:
                    # Carregar a imagem
                    img = Image.open(uploaded_file)

                    # Criar buffer para armazenar a nova imagem
                    img_io = io.BytesIO()

                    if formato_destino in ["JPEG", "JPG"]:
                        # Converter para RGB antes de salvar como JPG ou JPEG
                        img = img.convert("RGB")
                        img.save(img_io, "JPEG", quality=95)
                        mime_type = "image/jpeg"
                        file_extension = (
                            "jpg"  # Nome do arquivo será com extensão ".jpg"
                        )
                    elif formato_destino == "PNG":
                        img.save(img_io, "PNG")
                        mime_type = "image/png"
                        file_extension = "png"
                    elif formato_destino == "PDF":
                        img = img.convert("RGB")
                        img.save(img_io, "PDF")
                        mime_type = "application/pdf"
                        file_extension = "pdf"

                    # Redefinir o ponteiro para o início
                    img_io.seek(0)

                    st.success(
                        f"✅ Imagem convertida para {formato_destino.upper()} com sucesso!"
                    )

                    # Botão para download
                    st.download_button(
                        label=f"📥 Baixar {formato_destino.upper()}",
                        data=img_io,
                        file_name=f"{uploaded_file.name.split('.')[0]}_convertido.{file_extension}",
                        mime=mime_type,
                    )

                except Exception as e:
                    st.error(f"⚠️ Erro ao converter imagem: {e}")
        else:
            st.warning("⚠️ Por favor, envie um arquivo válido (Imagem ou PDF).")


# 🟢 FUNÇÃO "ORGANIZAÇÃO PLANILHA BANCÁRIA"
elif menu == "Organização Planilha Bancária":
    st.title("📑 Organização de Planilha Bancária")

    uploaded_file = st.file_uploader(
        "📂 Selecione uma planilha bancária", type=["xls", "xlsx"], key="bancaria"
    )

    if uploaded_file:
        st.session_state.uploaded_file_bancaria = uploaded_file

    # Se o arquivo foi enviado, processa
    if st.session_state.uploaded_file_bancaria:
        with st.spinner("Processando a planilha..."):
            output, df_processed = process_bank_statement(
                st.session_state.uploaded_file_bancaria
            )

            st.success("✅ Planilha processada com sucesso!")

            # Exibir a tabela processada
            st.write("### 📊 Dados Processados")
            st.dataframe(df_processed)

            # Calcular totais e diferença
            total_credito = df_processed.loc[
                df_processed["Historico"] == "TOTAL", "Valor Crédito"
            ].values[0]
            total_debito = df_processed.loc[
                df_processed["Historico"] == "TOTAL", "Valor Débito"
            ].values[0]
            diferenca = df_processed.loc[
                df_processed["Historico"].str.contains("DIFERENÇA", na=False),
                "Historico",
            ].values[0]

            # Exibir totais de forma visual
            st.write("### 📈 Resumo Financeiro")
            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric(label="💰 Total Crédito", value=total_credito)

            with col2:
                st.metric(label="📉 Total Débito", value=total_debito)

            with col3:
                st.metric(
                    label="🔍 Diferença (Crédito - Débito)",
                    value=diferenca.split(":")[-1].strip(),
                )

            # Disponibilizar o download da planilha processada
            st.download_button(
                label="📥 Baixar Planilha Processada",
                data=output,
                file_name="Planilha_Bancaria_Processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# 🟢 FUNÇÃO "CONTABILIDADE - EXTRATO ML"
elif menu == "Contabilidade - Extrato ML":
    st.title("📘 Contabilidade - Extrato Mercado Livre")

    uploaded_pdf = st.file_uploader("📂 Envie o arquivo PDF do extrato ML", type=["pdf"])

    if uploaded_pdf:
        try:
            texto = ""
            doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
            for page in doc:
                texto += page.get_text()

            linhas = texto.splitlines()
            transacoes = []
            bloco = ""
            padrao_data = re.compile(r"\d{2}-\d{2}-\d{4}")

            for linha in linhas:
                if padrao_data.match(linha.strip()):
                    if bloco:
                        transacoes.append(bloco.strip())
                    bloco = linha.strip()
                else:
                    bloco += " " + linha.strip()

            if bloco:
                transacoes.append(bloco.strip())

            dados_extraidos = []

            for transacao in transacoes:
                try:
                    data = re.search(r"\d{2}-\d{2}-\d{4}", transacao).group()
                    valor_raw = re.search(r"R\$ -?\d{1,3}(?:\.\d{3})*,\d{2}", transacao)
                    valor = valor_raw.group().replace("R$ ", "").replace(".", "").replace(",", ".") if valor_raw else ""

                    saldo_raw = re.findall(r"R\$ -?\d{1,3}(?:\.\d{3})*,\d{2}", transacao)
                    saldo = saldo_raw[-1].replace("R$ ", "").replace(".", "").replace(",", ".") if len(saldo_raw) > 1 else ""

                    id_match = re.findall(r"\b\d{9,}\b", transacao)
                    id_operacao = id_match[-1] if id_match else ""

                    descricao = re.sub(r"\d{2}-\d{2}-\d{4}", "", transacao)
                    descricao = re.sub(r"R\$ -?\d{1,3}(?:\.\d{3})*,\d{2}", "", descricao)
                    descricao = re.sub(r"\b\d{9,}\b", "", descricao)
                    descricao = descricao.strip()

                    dados_extraidos.append({
                        "Data": datetime.strptime(data, "%d-%m-%Y").date(),
                        "Descrição": descricao,
                        "ID da Operação": id_operacao,
                        "Valor": float(valor) if valor else "",
                        "Saldo": float(saldo) if saldo else "",
                    })
                except Exception as e:
                    st.warning(f"⚠️ Erro ao processar uma transação: {e}")

            df = pd.DataFrame(dados_extraidos)

            if not df.empty:
                st.success("✅ Transações extraídas com sucesso!")
                st.dataframe(df)

                # Download do Excel
                output = io.BytesIO()
                df.to_excel(output, index=False)
                output.seek(0)

                st.download_button(
                    label="📥 Baixar Excel",
                    data=output,
                    file_name="extrato_mercado_livre.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("Nenhuma transação encontrada no PDF.")
        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo PDF: {e}")
