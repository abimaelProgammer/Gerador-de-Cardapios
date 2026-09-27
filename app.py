import streamlit as st
import os
import importlib
from datetime import date
from pathlib import Path

import functions
importlib.reload(functions)
from functions import gerar_cardapio, gerar_cardapio_pdf

import etiquetas
importlib.reload(etiquetas)
from etiquetas import gerar_pdf_etiquetas

st.set_page_config(page_title="Lé Clair - Sistema de Impressão", page_icon="📋", layout="wide")

st.title("Lé Clair - Sistema de Impressão 📋")

# Criação das duas abas principais
aba_cardapio, aba_etiquetas = st.tabs(["📋 Gerador de Cardápio", "🏷️ Etiquetas Térmicas (Elgin L42)"])

# ==============================================================================
# ABA 1: GERADOR DE CARDÁPIO
# ==============================================================================
with aba_cardapio:
    st.write("Faça o upload da planilha de produtos para gerar o cardápio atualizado em Excel e PDF A4.")

    arquivo_produtos = st.file_uploader("Planilha de Produtos (Obrigatório, formato .xlsx)", type=["xlsx"], key="prod_cardapio")
    arquivo_modelo = st.file_uploader("Cardápio Anterior / Modelo (Opcional, formato .xlsx)", type=["xlsx"], key="mod_cardapio")

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

    if "planilha_gerada" not in st.session_state:
        st.session_state.planilha_gerada = None
    if "pdf_gerado" not in st.session_state:
        st.session_state.pdf_gerado = None

    if st.button("Gerar Cardápio", type="primary"):
        if arquivo_produtos is None:
            st.error("Por favor, faça o upload da planilha de produtos para continuar.")
        else:
            with st.spinner("Gerando cardápio em Excel e PDF..."):
                try:
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
                    
                    caminho_destino_xlsx = "Cardapio_Gerado.xlsx"
                    caminho_destino_pdf = "Cardapio_Gerado.pdf"
                    layout_escolhido = "paisagem" if "2 Vias" in formato_impressao else "retrato"
                    
                    # 1. Gera Planilha Excel
                    arquivo_gerado = gerar_cardapio(
                        fonte_produtos=caminho_produtos,
                        destino=caminho_destino_xlsx,
                        modelo_cardapio=caminho_modelo,
                        ocultar_pausados=not mostrar_pausados,
                        categorias_excluidas=categorias_excluidas,
                        layout=layout_escolhido,
                        itens_em_negrito=itens_em_negrito
                    )

                    # 2. Gera PDF Vetorial A4 nativo
                    arquivo_pdf = gerar_cardapio_pdf(
                        fonte_produtos=caminho_produtos,
                        destino=caminho_destino_pdf,
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
                    
                    st.success("Cardápio gerado com sucesso! Escolha o formato abaixo para baixar.")
                    
                except Exception as e:
                    st.error(f"Ocorreu um erro ao gerar o cardápio: {str(e)}")
                
                finally:
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

# ==============================================================================
# ABA 2: ETIQUETAS TÉRMICAS (ELGIN L42 PRO FULL - 100 x 65 mm)
# ==============================================================================
with aba_etiquetas:
    st.markdown("### Emissão de Etiquetas de Produtos (100 x 65 mm)")
    st.caption("Layout oficial Lé Clair com medidas exatas para impressora térmica Elgin L42 Pro Full.")

    # 1. Carrega a lista de produtos (da planilha subida no cardápio ou da planilha local)
    from functions import ler_produtos

    produtos_cadastrados = []
    if arquivo_produtos is not None:
        try:
            produtos_cadastrados = ler_produtos(arquivo_produtos)
        except Exception:
            pass
    elif Path("Produtos (58).xlsx").exists():
        try:
            produtos_cadastrados = ler_produtos("Produtos (58).xlsx")
        except Exception:
            pass

    # Apenas produtos ativos por padrão
    produtos_ativos = [p for p in produtos_cadastrados if str(p.get("status", "")).strip().lower() == "ativo"]

    # 2. Escolha entre selecionar da lista ou digitar manualmente
    if produtos_ativos:
        modo_origem = st.radio(
            "Origem do Produto:",
            options=["📋 Selecionar do Cardápio", "✍️ Digitar Manualmente"],
            horizontal=True,
            key="modo_origem_etiqueta"
        )
    else:
        modo_origem = "✍️ Digitar Manualmente"
        st.info("💡 Dica: Faça o upload da planilha na aba 'Gerador de Cardápio' para selecionar produtos automaticamente.")

    nome_sugerido = "Camarão empanado"
    modo_servir_sugerido = "Aqueça numa panela e sirva."
    conservacao_sugerida = 0  # 0: Refrigerado

    if modo_origem == "📋 Selecionar do Cardápio":
        col_filtro1, col_filtro2 = st.columns([1, 2])
        
        # Filtro opcional de categoria
        categorias_lista = ["Todas"] + sorted(list({str(p["categoria"]) for p in produtos_ativos}))
        with col_filtro1:
            cat_escolhida = st.selectbox("Filtrar por Categoria:", options=categorias_lista, key="filtro_cat_etiqueta")

        itens_filtrados = [
            p for p in produtos_ativos 
            if cat_escolhida == "Todas" or str(p["categoria"]) == cat_escolhida
        ]

        with col_filtro2:
            produto_selecionado = st.selectbox(
                "Pesquise ou selecione o Produto:",
                options=itens_filtrados,
                format_func=lambda p: f"{p['nome']} (Cód: {p['codigo']})",
                key="select_produto_etiqueta"
            )

        if produto_selecionado:
            nome_sugerido = str(produto_selecionado["nome"])
            
            # Sugestão inteligente de conservação dependendo da categoria
            cat_lower = str(produto_selecionado.get("categoria", "")).lower()
            if "bebida" in cat_lower or "vinho" in cat_lower or "cerveja" in cat_lower:
                modo_servir_sugerido = "Servir gelado."
                conservacao_sugerida = 2  # Local seco e fresco
            elif "destilado" in cat_lower or "aperitivo" in cat_lower:
                modo_servir_sugerido = "Servir em temperatura ambiente ou com gelo."
                conservacao_sugerida = 2
            else:
                modo_servir_sugerido = "Aqueça numa panela e sirva."
                conservacao_sugerida = 0  # Refrigerado

    st.markdown("#### Detalhes da Etiqueta")

    with st.form("form_etiqueta"):
        col_et1, col_et2 = st.columns(2)
        with col_et1:
            et_produto = st.text_input("Nome do Produto na Etiqueta", value=nome_sugerido)
            et_modo_servir = st.text_input("Modo de servir", value=modo_servir_sugerido)
            et_saudacao = st.text_input("Saudação", value="Bom apetite!")

        with col_et2:
            et_conservacao = st.selectbox(
                "Instrução de Armazenamento",
                options=["MANTER REFRIGERADO", "MANTER CONGELADO", "CONSERVAR EM LOCAL SECO E FRESCO"],
                index=conservacao_sugerida,
            )
            et_data_fab = st.date_input("Data de Fabricação", value=date.today(), format="DD/MM/YYYY")
            et_validade = st.text_input("Validade", value="3 dias refrigerado.")
            et_quantidade = st.number_input("Quantidade de etiquetas (cópias)", min_value=1, max_value=500, value=1)

        gerar_etiqueta_btn = st.form_submit_button("Gerar Etiquetas em PDF", type="primary")

    if gerar_etiqueta_btn:
        data_fab_formatada = et_data_fab.strftime("%d/%m/%Y")
        pdf_etiquetas_buffer = gerar_pdf_etiquetas(
            produto=et_produto,
            modo_servir=et_modo_servir,
            saudacao=et_saudacao,
            conservacao=et_conservacao,
            data_fabricacao=data_fab_formatada,
            validade=et_validade,
            quantidade=int(et_quantidade),
        )

        st.success(f"PDF com {et_quantidade} etiqueta(s) de '{et_produto}' gerado com sucesso!")
        st.download_button(
            label="📥 Baixar PDF das Etiquetas para a Elgin",
            data=pdf_etiquetas_buffer.getvalue(),
            file_name=f"Etiqueta_{et_produto.replace(' ', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )