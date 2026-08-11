# Gerador de Cardápio

Este é um aplicativo web feito com [Streamlit](https://streamlit.io/) que automatiza a geração de um cardápio em formato Excel (`.xlsx`) a partir de uma planilha de produtos. Ele utiliza o script de processamento interno (`functions.py`) para aplicar as formatações visuais, inserir a logo e organizar os itens em categorias.

## 🚀 Funcionalidades

- **Geração a partir de Planilha:** Faça o upload da sua planilha de produtos e receba o cardápio devidamente formatado.
- **Opção de Modelo Anterior:** É possível enviar um cardápio gerado anteriormente para aproveitar os avisos de rodapé, taxas, etc.
- **Filtros e Configurações:**
  - Escolha entre ocultar ou exibir produtos pausados/inativos.
  - Exclua determinadas categorias de produtos apenas digitando seus nomes.
- **Download Simplificado:** Após a geração, baixe a planilha final de forma rápida diretamente pela interface web.

## 🛠️ Tecnologias Utilizadas

- **Python 3**
- **Streamlit** (para a interface visual)
- **OpenPyXL** (para a leitura e estilização das planilhas Excel)

## 📦 Como Instalar e Rodar Localmente

### 1. Pré-requisitos
Certifique-se de ter o Python instalado na sua máquina.

### 2. Instalação das dependências
Recomenda-se o uso de um ambiente virtual. No seu terminal, na pasta do projeto, execute:

```bash
# Criação do ambiente virtual (opcional)
python -m venv .venv

# Instalação das dependências necessárias
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 3. Executando o aplicativo
Para rodar a interface web, digite o seguinte comando:
```bash
.venv\Scripts\python.exe -m streamlit run app.py
```
O aplicativo abrirá automaticamente no seu navegador no endereço padrão `http://localhost:8501`.

## 🌐 Como Hospedar (Streamlit Community Cloud)

O projeto está pronto para ser hospedado no [Streamlit Community Cloud](https://share.streamlit.io/):
1. Faça o upload ou sincronize estes arquivos (`app.py`, `functions.py`, `logo2.jpg`, e `requirements.txt`) em um repositório no **GitHub**.
2. Acesse o painel do Streamlit Cloud, clique em **New app**.
3. Selecione o repositório, branch e configure o campo *Main file path* como `app.py`.
4. Clique em **Deploy**! Em poucos minutos o site estará online.
