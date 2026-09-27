"""Módulo de geração de etiquetas térmicas em PDF para Elgin L42 Pro Full (100x65mm)."""

from io import BytesIO
from pathlib import Path
from PIL import Image as PILImage
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas


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