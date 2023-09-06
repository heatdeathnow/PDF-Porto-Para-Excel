from extract import get_info_from_page, get_info_from_qe, get_info_from_premio, get_info_from_totalizador, get_info_from_detalhamento
from concurrent.futures import ThreadPoolExecutor
from argparse import ArgumentParser
from time import perf_counter
from utils import fix_types
from threading import Lock
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
        pages = pdf.pages[1:-1]

        vars.qev = get_info_from_qe(pdf.pages[0])
        vars.premio = get_info_from_premio(pdf.pages[0])
        vars.total = get_info_from_totalizador(pdf.pages[0])
        vars.detalhe = get_info_from_detalhamento(pdf.pages[0])

        # Justificativa: Python automaticamente dá cpu_count + 4 workers para um ThreadPoolExecutor. Nesse programa há 2 ThreadPoolExecutors.
        with ThreadPoolExecutor(max_workers = vars.main_threads, thread_name_prefix = 'Thread_número') as tpx:
            lock = Lock()
            dfs = tpx.map(get_info_from_page, pages, [lock] * len(pages))

        df = pd.concat(dfs, ignore_index = True)

    fix_types(df)

    start_time = perf_counter()
    print(f'Formando arquivo Excel... ', end = '')

    writer = pd.ExcelWriter(output, 'openpyxl')
    df.to_excel(writer, 'Informação extraída', index = False)
    add_cols_extracted(writer.sheets['Informação extraída'])  # As duas colunas na primeira planilha

    writer.book.create_sheet('Quantidades e valores')
    add_cols_qe_valor(writer.sheets['Quantidades e valores'], writer.sheets['Informação extraída'].max_row)  # Segunda planilha com informações de inclusões, reativações, alterações, exclusões
    
    writer.book.create_sheet('Prêmio')
    add_premio(writer.sheets['Prêmio'], writer.sheets['Informação extraída'].max_row)  # Terceira planilha com informações de coparticipação etc
    
    writer.book.create_sheet('Totalizador')
    add_totalizador(writer.sheets['Totalizador'], writer.sheets['Informação extraída'].max_row)  # Quarta planilha com segurados e prêmios por plano
    
    writer.book.create_sheet('Detalhamento')
    add_detalhamento(writer.sheets['Detalhamento'], writer.sheets['Informação extraída'])

    format_sheets(writer.book)
    check_values(writer.book)
    print(f'arquivo Excel criado em {perf_counter() - start_time:.2f} segundos.')

    writer.close()
    print()
    print(f'{"-" * 67}\nTempo total de execução: {perf_counter() - runtime:.2f} segundos.')
