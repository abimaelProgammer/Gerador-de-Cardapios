import streamlit as st
import os
from pathlib import Path
from functions import gerar_cardapio

try:
    import win32com.client
except ImportError:
    win32com = None

st.set_page_config(page_title="Gerador de Cardápio", page_icon="📋")

st.title("Gerador de Cardápio 📋")
st.write("Faça o upload da planilha de produtos para gerar o cardápio atualizado.")

arquivo_produtos = st.file_uploader("Planilha de Produtos (Obrigatório, formato .xlsx)", type=["xlsx"])
arquivo_modelo = st.file_uploader("Cardápio Anterior / Modelo (Opcional, formato .xlsx)", type=["xlsx"])

mostrar_pausados = st.checkbox("Mostrar produtos pausados", value=False)
categorias_excluidas_str = st.text_input("Categorias excluídas (separadas por vírgula)", value="")

# Utilizamos o session_state do Streamlit para armazenar a planilha gerada na memória.
# Isso é necessário porque botões de download recarregam a página.
if "planilha_gerada" not in st.session_state:
    st.session_state.planilha_gerada = None

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
                
                arquivo_gerado = gerar_cardapio(
                    fonte_produtos=caminho_produtos,
                    destino=caminho_destino,
                    modelo_cardapio=caminho_modelo,
                    ocultar_pausados=not mostrar_pausados,
                    categorias_excluidas=categorias_excluidas
                )
                
                with open(arquivo_gerado, "rb") as f:
                    st.session_state.planilha_gerada = f.read()
                
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

if st.session_state.planilha_gerada:
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="📥 Baixar Cardápio",
            data=st.session_state.planilha_gerada,
            file_name="Cardapio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with col2:
        pode_gerar_pdf = win32com is not None
        if st.button(
            "🖨️ Abrir PDF para Imprimir",
            type="secondary",
            use_container_width=True,
            disabled=not pode_gerar_pdf,
        ):
            with st.spinner("Gerando PDF do Cardápio..."):
                try:
                    temp_print_path = os.path.abspath("temp_imprimir.xlsx")
                    temp_pdf_path = os.path.abspath("temp_imprimir.pdf")
                    
                    with open(temp_print_path, "wb") as f:
                        f.write(st.session_state.planilha_gerada)
                    
                    if os.path.exists(temp_pdf_path):
                        os.remove(temp_pdf_path)
                        
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    wb = excel.Workbooks.Open(temp_print_path)
                    # 0 é o código para ExportAsFixedFormat tipo PDF
                    wb.ActiveSheet.ExportAsFixedFormat(0, temp_pdf_path)
                    wb.Close(False)
                    excel.Quit()
                    
                    os.startfile(temp_pdf_path)
                    st.success("PDF aberto! Escolha a sua impressora no leitor de PDF.")
                except Exception as e:
                    st.error(f"Não foi possível gerar o PDF. Verifique se o Excel está instalado: {e}")

        if not pode_gerar_pdf:
            st.caption(
                "Para gerar o PDF, instale o suporte do Windows com: "
                "`python -m pip install pywin32`. O download em Excel continua disponível."
            )
