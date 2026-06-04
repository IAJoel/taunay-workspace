# -*- coding: utf-8 -*-
"""
gerar_matriz_direto.py
Gera MATRIZ_SENTENCAS_YYYY-MM-DD.xlsx com os dados analisados diretamente.
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, timedelta

PASTA = os.path.dirname(os.path.abspath(__file__))
HOJE  = date.today()
DATA_SITUACAO = HOJE - timedelta(days=2)

COR_HEADER = "1F3864"
COR_BRANCO = "FFFFFF"
COR_CINZA  = "F2F2F2"

BORDA = Border(
    left=Side(style="thin"),  right=Side(style="thin"),
    top=Side(style="thin"),   bottom=Side(style="thin"),
)

CABECALHOS = [
    "Número do processo",
    "Fase atual",
    "1ª Inst. Dano Moral",
    "1ª Inst. Dano Material",
    "Situação 1ª instância",
    "Data situação 1ª inst.",
    "Juiz 1ª instância",
    "Quantidade Dias Sem Energia",
    "Matéria",
]

LARGURAS = [34, 20, 20, 22, 36, 22, 44, 28, 44]

# Dados analisados: (CNJ, dano_moral, dano_material, situacao, juiz, dias_sem_energia, materia)
# dano_moral: valor em centavos (ex: 300000 = R$3.000,00) ou "000" se improcedente
# dano_material: valor em centavos ou 0 se improcedente/ilíquido
# dias_sem_energia: string com número ou "" se não aplicável
# materia: valor da coluna "Descrição dos Objetos > Pasta > Pasta Tarefa"
DADOS = [
    (
        "0890602-25.2025.8.19.0001",
        "000",
        "0",
        "Improcedente",
        "Lorena Reis Bastos Dutra",
        "",
        "INTERRUPÇÃO DO FORNECIMENTO DE ENERGIA ELÉTRICA",
    ),
    (
        "0806178-46.2025.8.19.0067",
        "000",
        "0",
        "Procedente",
        "Davi da Silva Grasso",
        "",
        "FATURAS",
    ),
    (
        "0804780-64.2025.8.19.0067",
        "000",
        "0",
        "Procedente",
        "Davi da Silva Grasso",
        "",
        "FATURAS",
    ),
    (
        "0807808-48.2024.8.19.0205",
        "000",
        "0",
        "Improcedente",
        "Ricardo Pinheiro Machado",
        "",
        "CADASTRO - ENCERRAMENTO DE CONTRATO",
    ),
    (
        "0806733-97.2024.8.19.0067",
        "000",
        "0",
        "Procedente",
        "Jeison Anders Tavares",
        "",
        "FATURAS - AUMENTO DE CONSUMO",
    ),
    (
        "0808101-10.2025.8.19.0067",
        "500000",
        "0",
        "Procedente",
        "Jeison Anders Tavares",
        "",
        "FATURAS",
    ),
    (
        "0837601-70.2022.8.19.0021",
        "000",
        "0",
        "Procedente",
        "Isabel Teresa Pinto Coelho Diniz",
        "",
        "FATURAS - AUMENTO DE CONSUMO",
    ),
    (
        "0819489-15.2024.8.19.0205",
        "000",
        "0",
        "Procedente",
        "Ricardo Pinheiro Machado",
        "",
        "Cobrança - Negativação SERASA/SPC",
    ),
    (
        "3000678-85.2026.8.19.0021",
        "600000",
        "0",
        "Procedente",
        "Não identificado",
        "",
        "INTERRUPÇÃO DO FORNECIMENTO DE ENERGIA ELÉTRICA",
    ),
]


def gravar_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MATRIZ"

    # Cabeçalho
    for col_i, texto in enumerate(CABECALHOS, 1):
        cel = ws.cell(row=1, column=col_i, value=texto)
        cel.font = Font(bold=True, color=COR_BRANCO, size=10)
        cel.fill = PatternFill("solid", fgColor=COR_HEADER)
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = BORDA
    ws.row_dimensions[1].height = 40

    # Dados
    for row_i, (cnj, dm, dmat, sit, juiz, dias, materia) in enumerate(DADOS, 2):
        fundo = COR_CINZA if row_i % 2 == 0 else COR_BRANCO
        valores = [
            cnj,
            "Sentença/Acórdão",
            dm,
            dmat,
            sit,
            DATA_SITUACAO,
            juiz,
            dias,
            materia,
        ]
        for col_i, valor in enumerate(valores, 1):
            cel = ws.cell(row=row_i, column=col_i, value=valor)
            cel.fill = PatternFill("solid", fgColor=fundo)
            cel.border = BORDA
            if col_i == 6 and isinstance(valor, date):
                cel.number_format = "DD/MM/YYYY"
                cel.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row_i].height = 28

    for col_i, larg in enumerate(LARGURAS, 1):
        ws.column_dimensions[get_column_letter(col_i)].width = larg

    saida = os.path.join(PASTA, f"MATRIZ_SENTENCAS_{HOJE.isoformat()}.xlsx")
    wb.save(saida)
    return saida


if __name__ == "__main__":
    saida = gravar_excel()
    print(f"Planilha gerada: {saida}")
    print(f"Linhas de dados: {len(DADOS)}")
    print(f"Data situação: {DATA_SITUACAO.strftime('%d/%m/%Y')}")
