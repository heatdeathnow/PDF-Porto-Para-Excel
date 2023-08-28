from extract import get_info_from_qe_valores, get_info_from_premio, get_info_from_totalizador, get_info_from_detalhamento, get_info_from_page
from argparse import ArgumentParser
from time import perf_counter
from utils import fix_types
from format import *
import pandas as pd
from sheet import *
import pdfplumber
import vars


if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument('input')
    args = parser.parse_args()
    file = args.input

    if file[-4:] != '.pdf':
        raise TypeError('Apenas arquivos .PDF são permitidos.')
    output = file.replace('.pdf', '.xlsx').replace('.PDF', '.xlsx')

    runtime = perf_counter()
    with pdfplumber.open(file) as pdf:
        amount = len(pdf.pages)

        vars.qev = get_info_from_qe_valores(pdf.pages[0])
        vars.premio = get_info_from_premio(pdf.pages[0])
        vars.total = get_info_from_totalizador(pdf.pages[0])
        vars.detalhe = get_info_from_detalhamento(pdf.pages[0])
        for i in range(1, amount - 1):  # Ignora a primeira e última página.
            start_time = perf_counter()
            print(f'Extraindo informação da página {i + 1:02}ª... ', end = '')
            df = get_info_from_page(pdf.pages[i])
            print(f'informação extraída em {perf_counter() - start_time:.2f} segundos.')

            if i == 1:
                baseframe = df.copy()
            else:
                baseframe = pd.concat([baseframe, df], ignore_index = True)

    fix_types(baseframe)

    start_time = perf_counter()
    print(f'Salvando a informação no disco... ', end = '')
    writer = pd.ExcelWriter(output, 'openpyxl')
    baseframe.to_excel(writer, 'Informação extraída', index = False)
    print(f'Informação salva em {perf_counter() - start_time:.2f} segundos.')

    start_time = perf_counter()
    print(f'Adicionando fórmulas do Excel... ', end = '')
    add_cols_extracted(writer.sheets['Informação extraída'])  # As duas colunas na primeira planilha
    writer.book.create_sheet('Quantidades e valores')
    add_cols_qe_valor(writer.sheets['Quantidades e valores'], writer.sheets['Informação extraída'].max_row)  # Segunda planilha com informações de inclusões, reativações, alterações, exclusões
    writer.book.create_sheet('Prêmio')
    add_premio(writer.sheets['Prêmio'], writer.sheets['Informação extraída'].max_row)  # Terceira planilha com informações de coparticipação etc
    writer.book.create_sheet('Totalizador')
    add_totalizador(writer.sheets['Totalizador'], writer.sheets['Informação extraída'].max_row)  # Quarta planilha com segurados e prêmios por plano
    writer.book.create_sheet('Detalhamento')
    add_detalhamento(writer.sheets['Detalhamento'], writer.sheets['Informação extraída'])
    print(f'fórmulas adicionadas em {perf_counter() - start_time:.2f} segundos.')

    start_time = perf_counter()
    print(f'Formatando planilhas... ', end = '')
    format_sheets(writer.book)
    check_values(writer.book)
    print(f'planilhas formatadas em {perf_counter() - start_time:.2f} segundos.')

    writer.close()
    print(f'Tempo total de execução: {perf_counter() - runtime:.2f} segundos.')
