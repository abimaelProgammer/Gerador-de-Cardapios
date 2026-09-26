"""Gera o cardapio a partir de uma planilha de produtos.

Uso pela linha de comando:
    python functions.py --produtos "Produtos (58).xlsx" --saida "Cardapio.xlsx" --layout paisagem

A fonte pode ser trocada tanto pelo argumento ``--produtos`` quanto ao chamar
diretamente a funcao ``gerar_cardapio``.
"""

from __future__ import annotations

import argparse
from collections import OrderedDict
from datetime import date
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


COLUNAS_OBRIGATORIAS = {
    "codigo": ("Cód. Sistema", "Cod. Sistema", "Código", "Codigo"),
    "categoria": ("Categoria",),
    "nome": ("Nome", "Produto", "Item"),
    "preco": ("Preço Venda", "Preco Venda", "Preço", "Preco"),
    "status": ("Status Venda", "Status"),
}

RODAPES_PADRAO = [
    ("Cobramos 10% como taxa de serviço", "Cobramos 10% como taxa de serviço"),
    (
        "Rede WI-FI: Clientes Le Clair    Senha: @Clientes2026",
        "Rede WI-FI: Clientes Le Clair    Senha: @Clientes2026",
    ),
    ("PIX (CNPJ) 64.111.665/0001-55", "PIX (CPF) 428.255.047-34"),
]


def _normalizar(texto: Any) -> str:
    import unicodedata

    texto = "" if texto is None else str(texto)
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    ).strip().casefold()


def _localizar_colunas(cabecalhos: Iterable[Any]) -> dict[str, int]:
    normalizados = {_normalizar(valor): indice for indice, valor in enumerate(cabecalhos, 1)}
    resultado: dict[str, int] = {}
    for campo, alternativas in COLUNAS_OBRIGATORIAS.items():
        for alternativa in alternativas:
            if _normalizar(alternativa) in normalizados:
                resultado[campo] = normalizados[_normalizar(alternativa)]
                break
        if campo not in resultado and campo != "status":
            raise ValueError(
                f"Coluna obrigatória não encontrada: {alternativas[0]}. "
                "Verifique o cabeçalho da planilha de produtos."
            )
    return resultado


def ler_produtos(fonte_produtos: str | Path, aba: str | None = None) -> list[dict[str, Any]]:
    """Lê produtos de qualquer .xlsx com os cabeçalhos esperados."""
    caminho = Path(fonte_produtos)
    if not caminho.exists():
        raise FileNotFoundError(f"Fonte de produtos não encontrada: {caminho}")

    wb = load_workbook(caminho, data_only=True, read_only=True)
    try:
        ws = wb[aba] if aba else wb[wb.sheetnames[0]]
        colunas = _localizar_colunas(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        produtos = []
        for numero_linha, linha in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            def valor(campo: str) -> Any:
                indice = colunas.get(campo)
                return linha[indice - 1] if indice and indice <= len(linha) else None

            if valor("codigo") is None or not valor("nome"):
                continue
            produtos.append(
                {
                    "codigo": valor("codigo"),
                    "categoria": valor("categoria") or "SEM CATEGORIA",
                    "nome": valor("nome"),
                    "preco": valor("preco") or 0,
                    "status": valor("status") or "Ativo",
                    "linha_origem": numero_linha,
                }
            )
        return produtos
    finally:
        wb.close()


def _rodapes_do_modelo(
    modelo_cardapio: str | Path | None,
) -> list[tuple[str, str]]:
    """Aproveita os avisos das duas metades do cardápio anterior."""
    if not modelo_cardapio or not Path(modelo_cardapio).exists():
        return list(RODAPES_PADRAO)
    wb = load_workbook(modelo_cardapio, data_only=False, read_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        textos: list[tuple[str, str]] = []
        palavras_chave = (
            "taxa", "wi-fi", "wifi", "pix", "servico", "gorjeta",
            "funcionamento", "reserva", "pedido minimo", "senha", "cnpj", "cpf",
            "pagamento", "cartao", "dinheiro",
        )

        def texto_do_bloco(linha: tuple[Any, ...], inicio: int, fim: int) -> str:
            for valor in linha[inicio:fim]:
                if valor is None:
                    continue
                texto = str(valor).strip()
                if texto and not texto.startswith("="):
                    return texto
            return ""

        for linha in ws.iter_rows(values_only=True):
            texto_esquerda = texto_do_bloco(linha, 1, 4)
            texto_direita = texto_do_bloco(linha, 4, 7)
            candidatos = (texto_esquerda, texto_direita)

            if not any(
                any(chave in _normalizar(texto) for chave in palavras_chave)
                for texto in candidatos if texto
            ):
                continue

            texto_esquerda = texto_esquerda or texto_direita
            texto_direita = texto_direita or texto_esquerda
            rodape = (texto_esquerda, texto_direita)
            if max(map(len, rodape)) > 10 and rodape not in textos:
                textos.append(rodape)
        return textos or list(RODAPES_PADRAO)
    finally:
        wb.close()


def _estilizar_faixa(ws, min_col: int, max_col: int, row: int, *, fill=None, font=None, alignment=None, border=None):
    """Aplica formatação uniforme em todas as células de uma faixa na mesma linha."""
    for col in range(min_col, max_col + 1):
        c = ws.cell(row, col)
        if fill is not None:
            c.fill = fill
        if font is not None:
            c.font = font
        if alignment is not None:
            c.alignment = alignment
        if border is not None:
            c.border = border


def gerar_cardapio(
    fonte_produtos: str | Path,
    destino: str | Path = "Cardapio.xlsx",
    modelo_cardapio: str | Path | None = "cardapio 2026-07-28.xlsx",
    *,
    aba_produtos: str | None = None,
    ocultar_pausados: bool = True,
    categorias_excluidas: Iterable[str] = (),
    layout: str = "paisagem",
    itens_em_negrito: bool = True,
) -> Path:
    """Cria um novo Excel de cardápio otimizado para folha A4 e retorna o caminho gerado.

    ``layout``:
        - "paisagem": 2 vias idênticas lado a lado (para corte ao meio da folha A4).
        - "retrato": 1 via ocupando a folha A4 inteira com fontes grandes e espaçadas.
    ``itens_em_negrito``:
        - Se True, destaca os nomes dos itens e códigos em negrito.
    """
    layout_norm = layout.strip().lower() if isinstance(layout, str) else "paisagem"
    is_paisagem = layout_norm == "paisagem"

    produtos = ler_produtos(fonte_produtos, aba_produtos)
    excluidas = {_normalizar(c) for c in categorias_excluidas}
    produtos = [p for p in produtos if _normalizar(p["categoria"]) not in excluidas]

    grupos: OrderedDict[str, list[dict[str, Any]]] = OrderedDict()
    for produto in produtos:
        grupos.setdefault(str(produto["categoria"]), []).append(produto)

    wb = Workbook()
    ws = wb.active
    ws.title = "Cardápio"
    apoio = wb.create_sheet("Configuração")

    # Paleta de cores oficial do cardápio
    azul = "1F4E78"
    azul_claro = "D9EAF7"
    branco = "FFFFFF"
    cinza = "D9E1F2"
    bege_creme = "EBE2D1"
    verde_escuro = "1B291E"
    cor_texto_escuro = "000000"

    borda_fina = Border(*(Side(style="thin", color="808080") for _ in range(4)))
    borda_corte = Border(right=Side(style="dashed", color="A0A0A0"))

    # Configuração de tipografia e alturas dependendo do layout
    if is_paisagem:
        # Modo 2 vias em Paisagem A4 (meia folha para corte) - preenche toda a folha
        tam_titulo = 15
        tam_cabecalho = 11.5
        tam_categoria = 11.5
        tam_produto = 11
        tam_inativo = 9.5
        tam_rodape = 10.5

        altura_linha_titulo = 40
        altura_linha_cabecalho = 20
        altura_linha_categoria = 18
        altura_linha_produto = 16.5
        altura_linha_rodape = 17
        altura_logo = 34
    else:
        # Modo 1 via em Retrato A4 (folha inteira, máxima legibilidade)
        tam_titulo = 18
        tam_cabecalho = 15
        tam_categoria = 14
        tam_produto = 13.5
        tam_inativo = 11
        tam_rodape = 12.5

        altura_linha_titulo = 55
        altura_linha_cabecalho = 24
        altura_linha_categoria = 22
        altura_linha_produto = 19.5
        altura_linha_rodape = 21
        altura_logo = 48

    # Estilos de Fontes
    fonte_titulo = Font(size=tam_titulo, bold=True, color=verde_escuro)
    fonte_cabecalho = Font(size=tam_cabecalho, bold=True, color=branco)
    fonte_categoria = Font(size=tam_categoria, bold=True, color=azul)
    fonte_cod = Font(size=tam_produto, bold=True, color=cor_texto_escuro)
    fonte_nome = Font(size=tam_produto, bold=itens_em_negrito, color=cor_texto_escuro)
    fonte_preco = Font(size=tam_produto, bold=True, color=cor_texto_escuro)
    fonte_inativo = Font(size=tam_inativo, color="808080", italic=True)
    fonte_rodape = Font(size=tam_rodape, bold=True, color=cor_texto_escuro)

    fill_titulo = PatternFill("solid", fgColor=bege_creme)
    fill_cabecalho = PatternFill("solid", fgColor=azul)
    fill_categoria = PatternFill("solid", fgColor=azul_claro)
    fill_rodape = PatternFill("solid", fgColor=cinza)

    # 1. LINHA 1: Cabeçalho com Título e Logo
    ws.row_dimensions[1].height = altura_linha_titulo

    # Via Esquerda: mescla B1:D1 (não inclui coluna A que é oculta)
    ws.merge_cells("B1:D1")
    ws["B1"] = f"CARDÁPIO {date.today():%d/%m/%Y}"
    _estilizar_faixa(
        ws, 2, 4, 1,
        fill=fill_titulo, font=fonte_titulo,
        alignment=Alignment(horizontal="center", vertical="center"),
        border=borda_fina
    )

    if is_paisagem:
        # Linha de corte central no topo
        ws.cell(1, 5).value = "✂"
        ws.cell(1, 5).font = Font(size=9, color="888888")
        ws.cell(1, 5).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(1, 5).border = borda_corte

        # Via Direita: mescla G1:I1 (não inclui coluna F que é oculta)
        ws.merge_cells("G1:I1")
        ws["G1"] = f"CARDÁPIO {date.today():%d/%m/%Y}"
        _estilizar_faixa(
            ws, 7, 9, 1,
            fill=fill_titulo, font=fonte_titulo,
            alignment=Alignment(horizontal="center", vertical="center"),
            border=borda_fina
        )

    # Inserção de Logo
    caminho_logo = Path(fonte_produtos).parent / "logo2.jpg"
    if not caminho_logo.exists():
        caminho_logo = Path(__file__).parent / "logo2.jpg"
    if not caminho_logo.exists():
        caminho_logo = Path("logo2.jpg")

    if caminho_logo.exists():
        try:
            img = Image(str(caminho_logo))
            ratio = altura_logo / img.height
            img.width = max(1, int(img.width * ratio))
            img.height = altura_logo
            ws.add_image(img, "B1")

            if is_paisagem:
                img2 = Image(str(caminho_logo))
                img2.width = img.width
                img2.height = img.height
                ws.add_image(img2, "G1")
        except Exception:
            pass  # Prossegue sem logo caso ocorra algum problema com a imagem

    # 2. LINHA 2: Cabeçalhos das Colunas
    ws.row_dimensions[2].height = altura_linha_cabecalho

    # Via Esquerda
    ws.cell(2, 1, "IMP").font = fonte_cabecalho
    ws.cell(2, 1).fill = fill_cabecalho
    ws.cell(2, 1).border = borda_fina

    cabecalhos_visiveis = ("CÓD", "ITENS", "R$")
    for offset, texto in enumerate(cabecalhos_visiveis, 2):
        celula = ws.cell(2, offset, texto)
        celula.font = fonte_cabecalho
        celula.fill = fill_cabecalho
        celula.alignment = Alignment(horizontal="center", vertical="center")
        celula.border = borda_fina

    if is_paisagem:
        ws.cell(2, 5).border = borda_corte

        # Via Direita
        ws.cell(2, 6, "IMP").font = fonte_cabecalho
        ws.cell(2, 6).fill = fill_cabecalho
        ws.cell(2, 6).border = borda_fina

        for offset, texto in enumerate(cabecalhos_visiveis, 7):
            celula = ws.cell(2, offset, texto)
            celula.font = fonte_cabecalho
            celula.fill = fill_cabecalho
            celula.alignment = Alignment(horizontal="center", vertical="center")
            celula.border = borda_fina

    # 3. ITENS E CATEGORIAS
    linha_atual = 3
    ativos = pausados = 0

    for categoria, itens in grupos.items():
        # Linha da Categoria
        ws.row_dimensions[linha_atual].height = altura_linha_categoria
        ws.cell(linha_atual, 1, "S")

        # Via Esquerda: mescla B:D
        ws.merge_cells(start_row=linha_atual, start_column=2, end_row=linha_atual, end_column=4)
        ws.cell(linha_atual, 2, categoria.upper())
        _estilizar_faixa(
            ws, 2, 4, linha_atual,
            fill=fill_categoria, font=fonte_categoria,
            alignment=Alignment(horizontal="center", vertical="center"),
            border=borda_fina
        )

        if is_paisagem:
            ws.cell(linha_atual, 5).border = borda_corte
            ws.cell(linha_atual, 6, "S")

            # Via Direita: mescla G:I
            ws.merge_cells(start_row=linha_atual, start_column=7, end_row=linha_atual, end_column=9)
            ws.cell(linha_atual, 7, categoria.upper())
            _estilizar_faixa(
                ws, 7, 9, linha_atual,
                fill=fill_categoria, font=fonte_categoria,
                alignment=Alignment(horizontal="center", vertical="center"),
                border=borda_fina
            )

        tem_ativo = any(_normalizar(produto["status"]) == "ativo" for produto in itens)
        if not tem_ativo and ocultar_pausados:
            ws.row_dimensions[linha_atual].hidden = True

        linha_atual += 1

        # Linhas de Produtos
        for produto in itens:
            ativo = _normalizar(produto["status"]) == "ativo"
            ws.row_dimensions[linha_atual].height = altura_linha_produto

            f_cod = fonte_cod if ativo else fonte_inativo
            f_nome = fonte_nome if ativo else fonte_inativo
            f_preco = fonte_preco if ativo else fonte_inativo

            # Via Esquerda
            ws.cell(linha_atual, 1, "S" if ativo else "N")

            c_cod = ws.cell(linha_atual, 2, produto["codigo"])
            c_cod.alignment = Alignment(horizontal="center", vertical="center")
            c_cod.font = f_cod
            c_cod.border = borda_fina

            c_nome = ws.cell(linha_atual, 3, produto["nome"])
            c_nome.alignment = Alignment(horizontal="left", vertical="center")
            c_nome.font = f_nome
            c_nome.border = borda_fina

            c_preco = ws.cell(linha_atual, 4, produto["preco"])
            c_preco.alignment = Alignment(horizontal="right", vertical="center")
            c_preco.number_format = '#,##0.00'
            c_preco.font = f_preco
            c_preco.border = borda_fina

            if is_paisagem:
                ws.cell(linha_atual, 5).border = borda_corte

                # Via Direita
                ws.cell(linha_atual, 6, "S" if ativo else "N")

                c_cod_d = ws.cell(linha_atual, 7, produto["codigo"])
                c_cod_d.alignment = Alignment(horizontal="center", vertical="center")
                c_cod_d.font = f_cod
                c_cod_d.border = borda_fina

                c_nome_d = ws.cell(linha_atual, 8, produto["nome"])
                c_nome_d.alignment = Alignment(horizontal="left", vertical="center")
                c_nome_d.font = f_nome
                c_nome_d.border = borda_fina

                c_preco_d = ws.cell(linha_atual, 9, produto["preco"])
                c_preco_d.alignment = Alignment(horizontal="right", vertical="center")
                c_preco_d.number_format = '#,##0.00'
                c_preco_d.font = f_preco
                c_preco_d.border = borda_fina

            if not ativo:
                pausados += 1
                ws.row_dimensions[linha_atual].hidden = ocultar_pausados
            else:
                ativos += 1

            linha_atual += 1

    # 4. RODAPÉS
    rodapes = _rodapes_do_modelo(modelo_cardapio)
    if rodapes:
        for texto_esquerda, texto_direita in rodapes:
            ws.row_dimensions[linha_atual].height = altura_linha_rodape

            # Via Esquerda: mescla B:D
            ws.merge_cells(start_row=linha_atual, start_column=2, end_row=linha_atual, end_column=4)
            ws.cell(linha_atual, 2, texto_esquerda)
            _estilizar_faixa(
                ws, 2, 4, linha_atual,
                fill=fill_rodape, font=fonte_rodape,
                alignment=Alignment(horizontal="center", vertical="center"),
                border=borda_fina
            )

            if is_paisagem:
                ws.cell(linha_atual, 5).border = borda_corte

                # Via Direita: mescla G:I
                ws.merge_cells(start_row=linha_atual, start_column=7, end_row=linha_atual, end_column=9)
                ws.cell(linha_atual, 7, texto_direita)
                _estilizar_faixa(
                    ws, 7, 9, linha_atual,
                    fill=fill_rodape, font=fonte_rodape,
                    alignment=Alignment(horizontal="center", vertical="center"),
                    border=borda_fina
                )

            linha_atual += 1

    # 5. CONFIGURAÇÃO DAS COLUNAS (Preenchimento total da folha A4)
    if is_paisagem:
        ws.column_dimensions["A"].width = 3
        ws.column_dimensions["A"].hidden = True
        ws.column_dimensions["B"].width = 8.5
        ws.column_dimensions["C"].width = 65.0
        ws.column_dimensions["D"].width = 14.5

        ws.column_dimensions["E"].width = 4.0

        ws.column_dimensions["F"].width = 3
        ws.column_dimensions["F"].hidden = True
        ws.column_dimensions["G"].width = 8.5
        ws.column_dimensions["H"].width = 65.0
        ws.column_dimensions["I"].width = 14.5

        ws.print_area = f"B1:I{linha_atual - 1}"
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        # Margens estreitas e equilibradas para preencher o papel A4 sem sobras
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.25
        ws.page_margins.bottom = 0.2
        ws.page_margins.header = 0.05
        ws.page_margins.footer = 0.05
    else:
        ws.column_dimensions["A"].width = 3
        ws.column_dimensions["A"].hidden = True
        ws.column_dimensions["B"].width = 11.0
        ws.column_dimensions["C"].width = 82.0
        ws.column_dimensions["D"].width = 18.0

        ws.print_area = f"B1:D{linha_atual - 1}"
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        # Margens equilibradas para folha A4 vertical
        ws.page_margins.left = 0.3
        ws.page_margins.right = 0.3
        ws.page_margins.top = 0.35
        ws.page_margins.bottom = 0.3
        ws.page_margins.header = 0.05
        ws.page_margins.footer = 0.05

    # 6. CONFIGURAÇÃO DE IMPRESSÃO E ENQUADRAMENTO NA PÁGINA A4
    ws.freeze_panes = "B3"
    ws.auto_filter.ref = f"A2:D{linha_atual - 1}"
    ws.sheet_view.showGridLines = False
    ws.print_title_rows = "1:2"

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True

    # 7. ABA DE APOIO/CONFIGURAÇÃO (Oculta para nunca ser impressa como folha extra)
    apoio.sheet_state = "hidden"
    configuracoes = (
        ("Configuração", "Valor"),
        ("Fonte de produtos", str(Path(fonte_produtos).resolve())),
        ("Modelo consultado", str(Path(modelo_cardapio).resolve()) if modelo_cardapio else ""),
        ("Layout de impressão", "2 Vias (Paisagem)" if is_paisagem else "1 Via (Retrato)"),
        ("Itens em negrito", "Sim" if itens_em_negrito else "Não"),
        ("Produtos incluídos", len(produtos)),
        ("Produtos ativos", ativos),
        ("Produtos pausados", pausados),
        ("Pausados ocultos", "Sim" if ocultar_pausados else "Não"),
    )
    for linha in configuracoes:
        apoio.append(linha)
    for celula in apoio[1]:
        celula.font = Font(bold=True, color=branco)
        celula.fill = PatternFill("solid", fgColor=azul)
    apoio.column_dimensions["A"].width = 24
    apoio.column_dimensions["B"].width = 90

    caminho_destino = Path(destino)
    caminho_destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_destino)
    return caminho_destino.resolve()


def _argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gera Cardapio.xlsx a partir de um Excel de produtos.")
    parser.add_argument("--produtos", default="Produtos (58).xlsx", help="Excel que será a fonte dos produtos")
    parser.add_argument("--modelo", default="cardapio 2026-07-28.xlsx", help="Cardápio anterior usado como referência")
    parser.add_argument("--saida", default="Cardapio.xlsx", help="Nome/caminho do novo Excel")
    parser.add_argument("--mostrar-pausados", action="store_true", help="Não oculta as linhas de produtos pausados")
    parser.add_argument("--layout", choices=["paisagem", "retrato"], default="paisagem", help="Formato de impressão A4")
    parser.add_argument("--sem-negrito", action="store_true", help="Não destacar itens em negrito")
    return parser.parse_args()


if __name__ == "__main__":
    args = _argumentos()
    gerado = gerar_cardapio(
        args.produtos,
        args.saida,
        args.modelo,
        ocultar_pausados=not args.mostrar_pausados,
        layout=args.layout,
        itens_em_negrito=not args.sem_negrito,
    )
    print(f"Cardápio gerado ({args.layout}): {gerado}")
