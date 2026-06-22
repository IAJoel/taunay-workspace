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
        "0027980-87.2019.8.19.0021",
        "300000",
        "0",
        "Procedente em Parte",
        "Não identificado",
        "",
        "PARCELAMENTO CAMPO - VALOR DIVERGENTE",
    ),
    (
        "0031514-68.2021.8.19.0021",
        "300000",
        "0",
        "Procedente em Parte",
        "Não identificado",
        "",
        "FATURAS - AUMENTO DE CONSUMO",
    ),
    (
        "0801451-52.2024.8.19.0205",
        "000",
        "0",
        "Improcedente",
        "Marcelo Mondego de Carvalho Lima",
        "",
        "CADASTRO - ENCERRAMENTO DE CONTRATO",
    ),
    (
        "0803174-43.2023.8.19.0205",
        "300000",
        "0",
        "Procedente",
        "Alex Quaresma Ravache",
        "",
        "Cobrança - Negativação SERASA/SPC",
    ),
    (
        "0804159-04.2024.8.19.0067",
        "500000",
        "0",
        "Procedente",
        "Jeison Anders Tavares",
        "",
        "FATURAS - DIVERGÊNCIA QUANTO A ITENS FATURADOS",
    ),
    (
        "0814518-06.2023.8.19.0210",
        "700000",
        "0",
        "Procedente",
        "Milena Angélica Drumond Morais Diz",
        "2",
        "INTERRUPÇÃO DO FORNECIMENTO DE ENERGIA ELÉTRICA",
    ),
    (
        "0838692-64.2023.8.19.0021",
        "500000",
        "0",
        "Procedente em Parte",
        "Danilo Nunes Cronemberger Miranda",
        "",
        "FATURAS - AUMENTO DE CONSUMO",
    ),
    (
        "0847140-89.2024.8.19.0021",
        "000",
        "0",
        "Improcedente",
        "Maria Daniella Binato de Castro",
        "",
        "Cobrança - Negativação SERASA/SPC",
    ),
    (
        "0912020-19.2025.8.19.0001",
        "000",
        "0",
        "Improcedente",
        "Luiz Antonio Valiera do Nascimento",
        "",
        "FATURAS",
    ),
    (
        "0015204-50.2022.8.19.0021",
        "000",
        "994080",
        "Procedente em Parte",
        "Não identificado",
        "",
        "COBRANÇA POR IRREGULARIDADE",
    ),
    (
        "0017815-73.2022.8.19.0021",
        "500000",
        "153982",
        "Procedente",
        "Não identificado",
        "",
        "COBRANÇA POR IRREGULARIDADE",
    ),
    (
        "0805011-35.2025.8.19.0021",
        "400000",
        "0",
        "Procedente em Parte",
        "Carolina Saud Coutinho",
        "7",
        "Interrupção n o fornecimento de energia elétrica",
    ),
    (
        "0806471-28.2023.8.19.0021",
        "000",
        "0",
        "Extinção sem julgamento de Mérito",
        "Isabela Pinheiro Guimarães",
        "",
        "FATURAS - AUMENTO DE CONSUMO",
    ),
    (
        "0816562-42.2025.8.19.0205",
        "000",
        "0",
        "Improcedente",
        "Não identificado",
        "",
        "COBRANÇA - NEGATIVAÇÃO",
    ),
    (
        "0961876-20.2023.8.19.0001",
        "000",
        "0",
        "Improcedente",
        "Não identificado",
        "",
        "FATURAS - DIVERGÊNCIA QUANTO A ITENS FATURADOS",
    ),
    (
        "0819221-92.2023.8.19.0205",
        "1000000",
        "0",
        "Procedente em Parte",
        "Eric Baracho Dore Fernandes",
        "",
        "CADASTRO - ENCERRAMENTO DE CONTRATO",
    ),
    (
        "0830183-77.2023.8.19.0205",
        "1000000",
        "0",
        "Procedente em Parte",
        "Eric Baracho Dore Fernandes",
        "",
        "FATURAS - AUMENTO DE CONSUMO",
    ),
    (
        "0833582-17.2023.8.19.0205",
        "000",
        "0",
        "Improcedente",
        "Eric Baracho Dore Fernandes",
        "",
        "Cobrança – Protesto",
    ),
    (
        "0840942-36.2024.8.19.0021",
        "000",
        "0",
        "Procedente em Parte",
        "Maria Daniella Binato de Castro",
        "",
        "FATURAS - AUMENTO DE CONSUMO",
    ),
    (
        "0841831-54.2023.8.19.0205",
        "1000000",
        "960000",
        "Procedente em Parte",
        "Eric Baracho Dore Fernandes",
        "6",
        "INTERRUPÇÃO DO FORNECIMENTO DE ENERGIA ELÉTRICA",
    ),
    (
        "0001797-67.2021.8.19.0067",
        "000",
        "0",
        "Improcedente",
        "Não identificado",
        "",
        "FATURAS - AUMENTO DE CONSUMO",
    ),
    (
        "0004756-57.2018.8.19.0021",
        "400000",
        "93940",
        "Procedente em Parte",
        "Não identificado",
        "",
        "COBRANÇA POR IRREGULARIDADE",
    ),
    (
        "0058377-03.2017.8.19.0021",
        "1000000",
        "0",
        "Procedente",
        "Não identificado",
        "",
        "COBRANÇA POR IRREGULARIDADE",
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

    nome = f"MATRIZ_SENTENCAS_{HOJE.isoformat()}.xlsx"
    saida = os.path.join(PASTA, nome)
    if os.path.exists(saida):
        try:
            os.rename(saida, saida)
        except PermissionError:
            nome = f"MATRIZ_SENTENCAS_{HOJE.isoformat()}_v2.xlsx"
            saida = os.path.join(PASTA, nome)
    wb.save(saida)
    return saida


if __name__ == "__main__":
    saida = gravar_excel()
    print(f"Planilha gerada: {saida}")
    print(f"Linhas de dados: {len(DADOS)}")
    print(f"Data situação: {DATA_SITUACAO.strftime('%d/%m/%Y')}")
