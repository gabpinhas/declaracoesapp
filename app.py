import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="Gerador de Declarações",
                   page_icon="📄", layout="centered")

# Código de acesso que só seu cliente sabe
CODIGO_CORRETO = "rottweilers"

st.markdown("""
    <style>
        .main { background-color: #f5f5f5; }
        .stTextInput>div>div>input { text-align: center; font-size: 20px; }
        .stButton>button { background-color: #4CAF50; color: white; font-size: 18px; padding: 10px 24px; }
        .stButton>button:hover { background-color: #45a049; }
        .stDownloadButton>button { background-color: #007BFF; color: white; font-size: 18px; padding: 10px 24px; }
        .stDownloadButton>button:hover { background-color: #0056b3; }
    </style>
""", unsafe_allow_html=True)

# Título
st.title("📜 Gerador de Declarações")

# Subtítulo com descrição
st.markdown("🔐 **Digite o código de acesso para continuar**")

# Entrada para o código de acesso
codigo_digitado = st.text_input("Código de acesso", type="password")


if codigo_digitado == CODIGO_CORRETO:
    st.success("Código correto! Você pode continuar.")

    # Upload do arquivo
    st.markdown("📂 **Envie a planilha Excel para gerar as declarações:**")
    uploaded_file = st.file_uploader("Envie a planilha Excel", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(
                uploaded_file, sheet_name="Respostas ao formulário 1")

            # Gerar arquivo com declarações
            def gerar_declaracoes(df):
                def gerar_texto(row):
                    # genero_masculino = row['Seu Estado Civil:'].strip().endswith(
                    #    "o") or row['Seu Estado Civil:'].strip().endswith("l")
                    # domiciliado = "domiciliado" if genero_masculino else "domiciliada"
                    # portador = "portador" if genero_masculino else "portadora"
                    # voluntario = "voluntário" if genero_masculino else "voluntária"

                    return (f"{row['Seu Nome Completo:']}, {row['Sua profissão:'].lower()}, {row['Seu Estado Civil:'].lower()}, "
                            f"portador(a) do RG nº {row['Seu Registro Geral (RG):']} {
                        row['Órgão que Emitiu seu RG (Exemplo: Detran - RJ, SSP-CE, SSP-PE, etc):']} "
                        f"e CPF sob o número {
                                row['Seu CPF (Por favor, adicione os pontos e confira se está com os 11 números, Ex: XXX.XXX.XXX-XX):']}, "
                        f"residente e domiciliado(a) à {row['Rua:']}, n° {
                                row['Número:']}, bairro {row['Bairro:']}, "
                        f"CEP: {row['CEP:']}, {
                                row['Cidade:']}, {row['Estado:']}, "
                        f"atua como {row['Seu Cargo Atual na EJ:']} na empresa júnior, voluntário(a) desde {row['Mês e Ano que foi efetivado na EJ (como membro efetivo):']}.")

                df["Declaração"] = df.apply(gerar_texto, axis=1)

                # Criar um arquivo .docx com as declarações
                doc = Document()
                for text in df["Declaração"]:
                    doc.add_paragraph(text)

                # Salvar o documento em memória
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                return buffer

            buffer = gerar_declaracoes(df)

            # Botão para download
            st.download_button(
                label="📥 Baixar Declarações",
                data=buffer,
                file_name="declaracoes.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"⚠ Erro ao processar a planilha: {str(e)}")

elif codigo_digitado:  # Se o código foi digitado, mas está errado
    st.error("❌ Código incorreto. Tente novamente.")
