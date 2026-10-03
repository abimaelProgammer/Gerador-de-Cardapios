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
from etiquetas import (
    gerar_pdf_etiquetas, 
    obter_catalogo_etiquetas, 
    salvar_template_etiqueta,
    excluir_produto_catalogo
)

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
            arquivo_produtos.seek(0)
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

    # 2. Catálogo Unificado
    catalogo = obter_catalogo_etiquetas(produtos_ativos)
    
    opcoes_selecao = ["➕ Criar Novo Produto / Personalizado"] + [p["produto"] for p in catalogo]
    
    col_sel, col_btn_del = st.columns([5, 1])
    with col_sel:
        produto_escolhido = st.selectbox(
            "Selecione um Produto para Imprimir ou Editar:",
            options=opcoes_selecao,
            key="select_produto_unificado"
        )
    
    with col_btn_del:
        st.write("") # espaçamento vertical para alinhar com o selectbox
        st.write("")
        if produto_escolhido != "➕ Criar Novo Produto / Personalizado":
            if st.button("🗑️ Remover", help=f"Ocultar/remover '{produto_escolhido}' do catálogo de etiquetas", use_container_width=True):
                excluir_produto_catalogo(produto_escolhido)
                st.toast(f"'{produto_escolhido}' removido da lista!")
                st.rerun()

    # Identifica os valores pré-carregados
    item_atual = None
    if produto_escolhido != "➕ Criar Novo Produto / Personalizado":
        item_atual = next((item for item in catalogo if item["produto"] == produto_escolhido), None)

    # Valores sugeridos padrão ou carregados do catálogo
    nome_sugerido = item_atual["produto"] if item_atual else ""
    modo_servir_sugerido = item_atual.get("modo_servir", "Aqueça numa panela e sirva.") if item_atual else "Aqueça numa panela e sirva."
    saudacao_sugerida = item_atual.get("saudacao", "Bom apetite!") if item_atual else "Bom apetite!"
    validade_sugerida = item_atual.get("validade", "3 dias refrigerado.") if item_atual else "3 dias refrigerado."
    
    conservacao_map = {
        "MANTER REFRIGERADO": 0,
        "MANTER CONGELADO": 1,
        "CONSERVAR EM LOCAL SECO E FRESCO": 2
    }
    cons_texto = item_atual.get("conservacao", "MANTER REFRIGERADO") if item_atual else "MANTER REFRIGERADO"
    conservacao_index = conservacao_map.get(cons_texto, 0)

    # Informações visuais do Formulário
    with st.container(border=True):
        st.markdown("##### ✏️ Configuração da Etiqueta")

        with st.form("form_etiqueta"):
            col_et1, col_et2 = st.columns(2)
            with col_et1:
                et_produto = st.text_input("Nome do Produto na Etiqueta", value=nome_sugerido, placeholder="Ex: Bolo de Cenoura com Chocolate")
                et_modo_servir = st.text_input("Modo de servir / Consumo", value=modo_servir_sugerido)
                et_saudacao = st.text_input("Saudação", value=saudacao_sugerida)

            with col_et2:
                et_conservacao = st.selectbox(
                    "Instrução de Armazenamento",
                    options=["MANTER REFRIGERADO", "MANTER CONGELADO", "CONSERVAR EM LOCAL SECO E FRESCO"],
                    index=conservacao_index,
                )
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    et_data_fab = st.date_input("Fabricação", value=date.today(), format="DD/MM/YYYY")
                with col_d2:
                    et_validade = st.text_input("Validade", value=validade_sugerida)

                et_quantidade = st.number_input("Cópias a imprimir", min_value=1, max_value=500, value=1)

            st.divider()
            col_check, col_btn = st.columns([2, 1])
            with col_check:
                salvar_no_catalogo = st.checkbox("💾 Salvar dados atualizados no catálogo", value=True)
            with col_btn:
                gerar_etiqueta_btn = st.form_submit_button("🖨️ Gerar PDF para Impressão", type="primary", use_container_width=True)

    if gerar_etiqueta_btn:
        if not et_produto.strip():
            st.warning("⚠️ Informe o nome do produto para gerar a etiqueta.")
        else:
            if salvar_no_catalogo:
                salvar_template_etiqueta({
                    "produto": et_produto.strip(),
                    "modo_servir": et_modo_servir.strip(),
                    "saudacao": et_saudacao.strip(),
                    "conservacao": et_conservacao,
                    "validade": et_validade.strip()
                })
                st.toast(f"✅ Configurações salvas para '{et_produto}'!")

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

            st.success("✅ Arquivo pronto para a Elgin L42!")
            st.download_button(
                label=f"📥 Baixar Etiqueta em PDF ({et_quantidade} cópias)",
                data=pdf_etiquetas_buffer,
                file_name=f"Etiqueta_{et_produto.replace(' ', '_')}.pdf",
                mime="application/pdf",
                type="secondary",
                use_container_width=True
            )