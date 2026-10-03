"""Módulo de geração de etiquetas térmicas em PDF para Elgin L42 Pro Full (100x65mm)."""

import json
from pathlib import Path
from io import BytesIO
from PIL import Image as PILImage
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

ARQUIVO_TEMPLATES_ETIQUETAS = Path("etiquetas_salvas.json")

def carregar_dados_etiquetas() -> dict:
    """Carrega o dicionário com templates e lista de itens ocultados/excluídos."""
    padrao = {"templates": [], "ocultos": []}
    if not ARQUIVO_TEMPLATES_ETIQUETAS.exists():
        return padrao
    try:
        with open(ARQUIVO_TEMPLATES_ETIQUETAS, "r", encoding="utf-8") as f:
            dados = json.load(f)
            # Compatibilidade caso o arquivo antes fosse apenas uma lista
            if isinstance(dados, list):
                return {"templates": dados, "ocultos": []}
            return dados
    except Exception:
        return padrao

def carregar_templates_etiquetas() -> list[dict]:
    """Carrega a lista de templates (apenas a chave 'templates')."""
    dados = carregar_dados_etiquetas()
    return dados.get("templates", [])

def salvar_dados_etiquetas(dados: dict) -> bool:
    """Salva os dados de templates e ocultos no JSON."""
    try:
        with open(ARQUIVO_TEMPLATES_ETIQUETAS, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False

def salvar_template_etiqueta(etiqueta_data: dict) -> bool:
    """
    Salva ou atualiza um produto no catálogo:
    - Atualiza os templates dentro da estrutura {'templates': [...], 'ocultos': [...]}.
    - Se o produto estava na lista de ocultos, remove ele de lá.
    """
    nome_alvo = etiqueta_data.get("produto", "").strip().lower()
    if not nome_alvo:
        return False
    dados = carregar_dados_etiquetas()
    templates = dados.get("templates", [])
    # 1. Atualiza se já existir ou adiciona novo
    atualizado = False
    for i, item in enumerate(templates):
        if isinstance(item, dict) and item.get("produto", "").strip().lower() == nome_alvo:
            templates[i] = etiqueta_data
            atualizado = True
            break
    if not atualizado:
        templates.append(etiqueta_data)
    dados["templates"] = templates
    # 2. Se estava na lista de ocultos, desoculta!
    dados["ocultos"] = [
        o for o in dados.get("ocultos", []) 
        if o.strip().lower() != nome_alvo
    ]
    return salvar_dados_etiquetas(dados)

def excluir_produto_catalogo(nome_produto: str) -> bool:
    """
    Remove o produto dos templates E adiciona à lista de ocultos,
    garantindo que não reapareça mesmo se estiver na planilha.
    """
    dados = carregar_dados_etiquetas()
    nome_norm = nome_produto.strip().lower()

    # Remove dos templates
    dados["templates"] = [
        t for t in dados.get("templates", []) 
        if t.get("produto", "").strip().lower() != nome_norm
    ]

    # Adiciona aos ocultos (se já não estiver)
    ocultos = [o.lower() for o in dados.get("ocultos", [])]
    if nome_norm not in ocultos:
        dados.setdefault("ocultos", []).append(nome_produto.strip())

    return salvar_dados_etiquetas(dados)

def obter_catalogo_etiquetas(produtos_cardapio: list[dict] = None) -> list[dict]:
    """Retorna o catálogo unificado, ignorando itens que foram ocultados/excluídos."""
    dados = carregar_dados_etiquetas()
    salvos = dados.get("templates", [])
    ocultos_set = {o.strip().lower() for o in dados.get("ocultos", [])}

    # Só entram salvos que não estão na lista de ocultos
    catalogo = [s for s in salvos if s.get("produto", "").strip().lower() not in ocultos_set]
    mapa_existentes = {s["produto"].strip().lower() for s in catalogo}

    if produtos_cardapio:
        for p in produtos_cardapio:
            nome = str(p.get("nome", "")).strip()
            if not nome:
                continue

            nome_lower = nome.lower()
            # Ignora se foi ocultado pelo usuário OU se já está no catálogo
            if nome_lower in ocultos_set or nome_lower in mapa_existentes:
                continue

            cat_lower = str(p.get("categoria", "")).lower()
            if any(beb in cat_lower for beb in ["bebida", "vinho", "cerveja"]):
                modo_servir = "Servir gelado."
                conservacao = "CONSERVAR EM LOCAL SECO E FRESCO"
            elif any(dest in cat_lower for dest in ["destilado", "aperitivo"]):
                modo_servir = "Servir em temperatura ambiente ou com gelo."
                conservacao = "CONSERVAR EM LOCAL SECO E FRESCO"
            else:
                modo_servir = "Aqueça numa panela e sirva."
                conservacao = "MANTER REFRIGERADO"

            catalogo.append({
                "produto": nome,
                "modo_servir": modo_servir,
                "saudacao": "Bom apetite!",
                "conservacao": conservacao,
                "validade": "3 dias refrigerado."
            })
            mapa_existentes.add(nome_lower)

    catalogo.sort(key=lambda x: x.get("produto", "").lower())
    return catalogo

def _obter_logo_recortada() -> str | None:
    """Prepara a logo com fundo transparente recortado para máxima nitidez."""
    caminho_png = Path("logo.png")
    caminho_jpg = Path("logo2.jpg")

    if caminho_png.exists():
        try:
            im = PILImage.open(caminho_png)
            bbox = im.getbbox()
            if bbox:
                im_cropped = im.crop(bbox)
                temp_logo = Path(".temp_logo_etiqueta.png")
                im_cropped.save(temp_logo, format="PNG")
                return str(temp_logo)
            return str(caminho_png)
        except Exception:
            return str(caminho_png)

    if caminho_jpg.exists():
        return str(caminho_jpg)

    return None

def desenhar_uma_etiqueta(
    c: canvas.Canvas,
    produto: str,
    modo_servir: str,
    saudacao: str = "Bom apetite!",
    conservacao: str = "MANTER REFRIGERADO",
    data_fabricacao: str = "",
    validade: str = "3 dias refrigerado.",
):
    """Desenha uma única etiqueta de 100mm x 65mm nas coordenadas exatas."""
    largura = 100 * mm
    altura = 65 * mm

    # 1. Logo Circular Lé Clair (à esquerda)
    caminho_logo = _obter_logo_recortada()
    if caminho_logo:
        try:
                       # Diâmetro reduzido (17mm) na parte superior esquerda
            c.drawImage(
                caminho_logo,
                5 * mm,         # x: encostado à esquerda com 5mm de margem
                43 * mm,        # y: parte superior da etiqueta
                width=17 * mm,  # largura reduzida
                height=17 * mm, # altura reduzida
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    # 2. Título do Produto (com linha sublinhada)
    tam_fonte_titulo = 15 if len(produto) <= 24 else 12
    c.setFont("Times-Bold", tam_fonte_titulo)
    c.drawCentredString(56 * mm, 54 * mm, produto)

    largura_texto = c.stringWidth(produto, "Times-Bold", tam_fonte_titulo)
    c.setLineWidth(1.0)
    c.line(56 * mm - largura_texto / 2, 52.5 * mm, 56 * mm + largura_texto / 2, 52.5 * mm)

    largura_texto = c.stringWidth(produto, "Times-Bold", tam_fonte_titulo)
    c.setLineWidth(1.0)
    c.line(50 * mm - largura_texto / 2, 52.5 * mm, 50 * mm + largura_texto / 2, 52.5 * mm)

    # 3. Modo de servir e Saudação
    c.setFont("Times-Roman", 10.5)
    if len(modo_servir) > 35:
        palavras = modo_servir.split()
        meio = len(palavras) // 2
        linha1 = " ".join(palavras[:meio])
        linha2 = " ".join(palavras[meio:])
        c.drawCentredString(56 * mm, 46.5 * mm, f"Modo de servir: {linha1}")
        c.drawCentredString(56 * mm, 42.5 * mm, linha2)
        c.drawCentredString(56 * mm, 38.0 * mm, saudacao)
    else:
        c.drawCentredString(56 * mm, 45.0 * mm, f"Modo de servir: {modo_servir}")
        c.drawCentredString(56 * mm, 40.5 * mm, saudacao)

    # 4. Alerta de Conservação
    c.setFont("Times-Bold", 13)
    partes_conservacao = conservacao.upper().split()
    if len(partes_conservacao) == 2:
        c.drawCentredString(50 * mm, 30.5 * mm, partes_conservacao[0])
        c.drawCentredString(50 * mm, 25.5 * mm, partes_conservacao[1])
    else:
        c.drawCentredString(50 * mm, 28.0 * mm, conservacao.upper())

    # 5. Datas (Fabricação e Validade)
    c.setFont("Times-Roman", 11.5)
    c.drawCentredString(50 * mm, 19.5 * mm, f"Fabricação: {data_fabricacao}")
    c.drawCentredString(50 * mm, 14.5 * mm, f"Validade: {validade}")

    # 6. Rodapé Legal / Institucional
    c.setFont("Times-Bold", 6.8)
    c.drawCentredString(50 * mm, 9.5 * mm, "Lé Clair Gastronomia & Cultura")

    c.setFont("Times-Roman", 6.2)
    c.drawCentredString(50 * mm, 7.3 * mm, "CNPJ: 64.111.665/0001-55")
    c.drawCentredString(50 * mm, 5.1 * mm, "Av. Barão de Studart , 1420, loja 08 - Aldeota - Fortaleza - CE")
    c.drawCentredString(50 * mm, 2.9 * mm, "Fone: (85) 99113-9073")

def gerar_pdf_etiquetas(
    produto: str,
    modo_servir: str,
    data_fabricacao: str,
    validade: str,
    conservacao: str = "MANTER REFRIGERADO",
    saudacao: str = "Bom apetite!",
    quantidade: int = 1,
) -> BytesIO:
    """Gera um PDF em memória contendo N cópias da etiqueta de 100x65mm."""
    buffer = BytesIO()
    largura = 100 * mm
    altura = 65 * mm

    c = canvas.Canvas(buffer, pagesize=(largura, altura))

    for _ in range(max(1, quantidade)):
        desenhar_uma_etiqueta(
            c,
            produto=produto,
            modo_servir=modo_servir,
            saudacao=saudacao,
            conservacao=conservacao,
            data_fabricacao=data_fabricacao,
            validade=validade,
        )
        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer

if __name__ == "__main__":
    # Teste rápido se executado diretamente
    pdf = gerar_pdf_etiquetas(
        produto="Camarão empanado",
        modo_servir="Aqueça numa panela e sirva.",
        data_fabricacao="28/08/2026",
        validade="3 dias refrigerado.",
        quantidade=1,
    )
    with open("etiqueta_teste.pdf", "wb") as f:
        f.write(pdf.getvalue())
    print("Etiqueta teste gerada com sucesso: etiqueta_teste.pdf")