import streamlit as st
import os
import importlib
from pathlib import Path

import functions
importlib.reload(functions)
from functions import gerar_cardapio, gerar_cardapio_pdf

st.set_page_config(page_title="Gerador de Cardápio", page_icon="📋")

st.title("Gerador de Cardápio 📋")
st.write("Faça o upload da planilha de produtos para gerar o cardápio atualizado.")

arquivo_produtos = st.file_uploader("Planilha de Produtos (Obrigatório, formato .xlsx)", type=["xlsx"])
arquivo_modelo = st.file_uploader("Cardápio Anterior / Modelo (Opcional, formato .xlsx)", type=["xlsx"])

st.markdown("### Configurações de Impressão")
formato_impressao = st.radio(
    "Formato de Impressão na Folha A4:",
    options=[
        "2 Vias por Folha A4 (Paisagem - Meia folha para corte)",
        "Folha Inteira A4 (1 Via - Retrato)"
    ],
    index=0,
    help="• 2 Vias (Paisagem): Imprime 2 cardápios idênticos lado a lado com linha tracejada de corte ao meio (ideal para mesas).\n• 1 Via (Retrato): Cardápio único com letras grandes preenchendo toda a folha A4."
)

col_opcoes1, col_opcoes2 = st.columns(2)
with col_opcoes1:
    itens_em_negrito = st.checkbox("Destacar itens e preços em negrito", value=True)
with col_opcoes2:
    mostrar_pausados = st.checkbox("Mostrar produtos pausados", value=False)

categorias_excluidas_str = st.text_input("Categorias excluídas (separadas por vírgula)", value="")

# Utilizamos o session_state do Streamlit para armazenar a planilha gerada na memória.
# Isso é necessário porque botões de download recarregam a página.
if "planilha_gerada" not in st.session_state:
    st.session_state.planilha_gerada = None
if "pdf_gerado" not in st.session_state:
    st.session_state.pdf_gerado = None

if st.button("Gerar Cardápio", type="primary"):
    if arquivo_produtos is None:
        st.error("Por favor, faça o upload da planilha de produtos para continuar.")
    else:
        with st.spinner("Gerando cardápio..."):
            try:
                # Salvando temporariamente no diretório atual para que functions.py encontre a logo (logo2.jpg)
                caminho_produtos = "temp_produtos_upload.xlsx"
                with open(caminho_produtos, "wb") as f:
                    f.write(arquivo_produtos.getvalue())
                
                caminho_modelo = None
                if arquivo_modelo is not None:
                    caminho_modelo = "temp_modelo_upload.xlsx"
                    with open(caminho_modelo, "wb") as f:
                        f.write(arquivo_modelo.getvalue())
                
                categorias_excluidas = []
                if categorias_excluidas_str.strip():
                    categorias_excluidas = [cat.strip() for cat in categorias_excluidas_str.split(",")]
                
                caminho_destino = "Cardapio_Gerado.xlsx"
                layout_escolhido = "paisagem" if "2 Vias" in formato_impressao else "retrato"
                
                arquivo_gerado = gerar_cardapio(
                    fonte_produtos=caminho_produtos,
                    destino=caminho_destino,
                    modelo_cardapio=caminho_modelo,
                    ocultar_pausados=not mostrar_pausados,
                    categorias_excluidas=categorias_excluidas,
                    layout=layout_escolhido,
                    itens_em_negrito=itens_em_negrito
                )

                arquivo_pdf = gerar_cardapio_pdf(
                    fonte_produtos=caminho_produtos,
                    destino="Cardapio_Gerado.pdf",
                    modelo_cardapio=caminho_modelo,
                    ocultar_pausados=not mostrar_pausados,
                    categorias_excluidas=categorias_excluidas,
                    layout=layout_escolhido,
                    itens_em_negrito=itens_em_negrito
                )

                with open(arquivo_gerado, "rb") as f:
                    st.session_state.planilha_gerada = f.read()

                with open(arquivo_pdf, "rb") as f:
                    st.session_state.pdf_gerado = f.read()
                
                st.success("Cardápio gerado com sucesso! Clique no botão abaixo para baixar.")
                
            except Exception as e:
                st.error(f"Ocorreu um erro ao gerar o cardápio: {str(e)}")
            finally:
                # Limpeza de arquivos temporários
                if os.path.exists("temp_produtos_upload.xlsx"):
                    os.remove("temp_produtos_upload.xlsx")
                if caminho_modelo and os.path.exists(caminho_modelo):
                    os.remove(caminho_modelo)
                if os.path.exists("Cardapio_Gerado.xlsx"):
                    os.remove("Cardapio_Gerado.xlsx")
                if os.path.exists("Cardapio_Gerado.pdf"):
                    os.remove("Cardapio_Gerado.pdf")

if st.session_state.planilha_gerada:
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="📥 Baixar Cardápio (Excel)",
            data=st.session_state.planilha_gerada,
            file_name="Cardapio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with col2:
        st.download_button(
            label="🖨️ Baixar Cardápio em PDF (Pronto para Imprimir)",
            data=st.session_state.pdf_gerado,
            file_name="Cardapio.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )