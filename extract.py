from openpyxl.utils import get_column_letter
from pdfplumber.page import Page
from pdfplumber.pdf import PDF
from utils import update_sub
from pandas import DataFrame
import pandas as pd
from crop import *
import vars


def get_cells(page: Page, rel_pos: list[int]) -> list:
    lines = page.extract_text_lines(return_chars = False)
    cells = [''] * len(rel_pos)
    flag = False

    for i, line in enumerate(lines):
        if i + 1 == len(lines) and not flag:  # Se é o último.
            index = rel_pos.index(round(line['top']))
            cells[index] = line['text']

        elif flag:
            flag = False
        
        elif abs(line['top'] - lines[i + 1]['top']) < 10:  # Se estiver muito pergunto verticalmente, considerar como duas linhas do mesmo texto
            index = rel_pos.index(round(line['top']))
            cells[index] = '\n'.join([line['text'], lines[i + 1]['text']])
            flag = True

        elif not flag:
            index = rel_pos.index(round(line['top']))
            cells[index] = line['text']

    return cells

def get_relative_positions(page: Page) -> list:  # Passar idade pois é mais fácil
    positions = []
    lines = page.extract_text_lines(return_chars = False)
    for line in lines:
        positions.append(round(line['top']))
    return positions

def get_info_from_leftside(page: Page) -> DataFrame:
    rel_pos = get_relative_positions(crop_idade(page))
    seguro = get_cells(crop_seguro(page), rel_pos)
    dep = get_cells(crop_dep(page), rel_pos)
    nome = get_cells(crop_nome(page), rel_pos)
    regfunc = get_cells(crop_regfunc(page), rel_pos)
    idade = get_cells(crop_idade(page), rel_pos)
    parentesco = get_cells(crop_parentesco(page), rel_pos)
    plano = get_cells(crop_plano(page), rel_pos)
    mov = get_cells(crop_mov(page), rel_pos)
    situacao = get_cells(crop_situacao(page), rel_pos)

    df = DataFrame({'Seguro': seguro,
                    'Dep': dep,
                    'Nome do Segurado': nome,
                    'Reg. Func.': regfunc,
                    'Idade': idade,
                    'Parentesco': parentesco,
                    'Plano': plano,
                    'Mov./Inic.Vig.': mov,
                    'Situação do Segurado': situacao})
    return df

def get_info_from_rightside(page: Page) -> DataFrame:
    rightside = crop_right_side(page)

    avoid_error = max(rightside.extract_text().count('Prêmio Base'), rightside.extract_text().count('Desc por Co-Part'))

    premio = [''] * avoid_error
    rataq = [''] * avoid_error
    ratav = [''] * avoid_error
    consultasq = [''] * avoid_error
    consultasv = [''] * avoid_error
    examesq = [''] * avoid_error
    examesv = [''] * avoid_error
    socorroq =[''] * avoid_error
    socorrov = [''] * avoid_error
    copart = [''] * avoid_error
    bonificacao = [''] * avoid_error
    acerto = [''] * avoid_error
    iof = [''] * avoid_error

    i = -1  # Prêmio Base vai ser sempre o primeiro valor.
    for line in rightside.extract_text_lines(return_chars = False):
        if 'Prêmio Base' in line['text']:
            premio[i + 1] = line['text'][12:].strip()
            i += 1

        elif 'CONSULTAS' in line['text']:
            end = line['text'].index(')')
            consultasq[i] = line['text'][11:end]  # O "(" sempre estará na 11ª posição.
            consultasv[i] = line['text'][end + 1:].strip()  # Começando depois do ")" retire todos os espaços.

        elif 'EXAMES' in line['text']:
            end = line['text'].index(')')
            examesq[i] = line['text'][8:end]  # O "(" sempre estará na 8ª posição.
            examesv[i] = line['text'][end + 1:].strip()  # Começando depois do ")" retire todos os espaços.

        elif 'PRONTO-SOCORRO' in line['text']:
            end = line['text'].index(')')
            socorroq[i] = line['text'][16:end]  # O "(" sempre estará na 16ª posição.
            socorrov[i] = line['text'][end + 1:].strip()  # Começando depois do ")" retire todos os espaços.

        elif 'Pro-Rata' in line['text']:
            end = line['text'].index(')')
            end_ = line['text'].index(')', end + 1)
            rataq[i] = line['text'][10:end]  # O "(" sempre estará na 16ª posição.
            ratav[i] = line['text'][end_ + 1:].strip()  # Começando depois do ")" retire todos os espaços.

        elif 'Desc por Co-Part' in line['text']:
            if copart[i] != '':  # Pode acontecer de faltar o prêmio por erro da Porto Seguro.
                i += 1
            copart[i] = line['text'][22:].strip()
        
        elif 'Bonificação GRES(-)' in line['text']:
            bonificacao[i] = line['text'][19:].strip()
        
        elif 'Acerto (-)' in line['text']:
            acerto[i] = line['text'][10:].strip()
        
        elif 'IOF' in line['text']:
            if iof[i] != '' and iof[i] != '0,00' and iof[i] != 0:
                pass  # Por algum motivo, os PDFs da Porto podem vir com dois IOFs, o segundo erroneamente zerado.
            else:
                iof[i] = line['text'][4:].strip()
        
        elif 'TOTAL' not in line['text'].upper():
            print(f"ERRO: Não foi possível ler a linha: {line['text']}")
            pass

    df = DataFrame({'Prêmio Base': premio,
                    'Pro-Rata (quantidade)': rataq,
                    'Pro-Rata (valor)': ratav,
                    'CONSULTAS (quantidade)': consultasq,
                    'CONSULTAS (valor)': consultasv,
                    'EXAMES (quantidade)': examesq,
                    'EXAMES (valor)': examesv,
                    'PRONTO-SOCORRO (quantidade)': socorroq,
                    'PRONTO-SOCORRO (valor)': socorrov,
                    'Desc por Co-Part. (-)': copart,
                    'Bonificação GRES(-)': bonificacao,
                    'Acerto (-)': acerto,
                    'IOF': iof})
    return df

def get_info_from_qe_valores(page: Page) -> DataFrame:
    cells = {}

    qev = crop_qe_valores(page)
    for i, row in enumerate(qev.extract_text_lines()):
        if i == 1: continue

        for j, value in enumerate(row['text'].split(' ')):
            if j % 2 == 0:
                cells[f'{get_column_letter(j + 2)}{i + 2}'] = int(value.replace('.', '').replace(',', '.').strip())  # B2:G2 & B4:G6
            else:
                cells[f'{get_column_letter(j + 2)}{i + 2}'] = float(value.replace('.', '').replace(',', '.').strip())
    return cells

def get_info_from_premio(page: Page) -> DataFrame:
    cells = {}
    premio = crop_premio(page)

    for i, row in enumerate(premio.extract_text_lines()):
        if i == 3: continue

        for j, value in enumerate(row['text'].split(' ')):
            cells[f'{get_column_letter(j + 2)}{i + 2}'] = float(value.replace('.', '').replace(',', '.').strip())
    return cells

def get_info_from_totalizador(page: Page) -> DataFrame:
    total = crop_totalizador(page)
    cells = {}

    for i, row in enumerate(total.extract_text_lines()):
        j = row['text'].index(':') + 2
        cells[f'B{i + 1}'] = float(row['text'][j:].replace('.', '').replace(',', '.').strip())
    return cells

def get_info_from_detalhamento(page: Page) -> DataFrame:
    detalhe = crop_detalhamento(page)
    cells = {}

    for row in detalhe.extract_text_lines():
        if 'Em atendimento à Lei' in row['text']: break
        l = 0
        for i in 'qwertyuiopasdfghjklçzxcvbnm':  # Isso descobre a posição da última letra na linha.
            j = row['text'].lower().rfind(i)
            if j > l: l = j

        hold = []
        for i, cell in enumerate(row['text'][l + 2:].split(' ')):
            if i == 3: continue

            if ',' in cell:
                hold.append(float(cell.replace('.', '').replace(',', '.').strip()))
            else:
                hold.append(int(cell.strip()))
        cells[row['text'][:l + 1]] = hold
    return cells

def get_info_from_page(page: Page) -> DataFrame:
    if vars.sub == '' or vars.sub_has_changed:
        update_sub(page)
        vars.sub_has_changed = False
        print(f'\n\n--- Subestipulante atualmente em leitura: {vars.sub} ---\n')

    check = crop_headers(page)
    if len(check.extract_words()) == 0:
        vars.sub_has_changed = True
        return DataFrame({'Seguro': [],
                          'Dep': [],
                          'Nome do Segurado': [],
                          'Reg. Func.': [],
                          'Idade': [],
                          'Parentesco': [],
                          'Plano': [],
                          'Mov./Inic.Vig.': [],
                          'Situação do Segurado': [],
                          'Prêmio Base': [],
                          'Pro-Rata (quantidade)': [],
                          'Pro-Rata (valor)': [],
                          'CONSULTAS (quantidade)': [],
                          'CONSULTAS (valor)': [],
                          'EXAMES (quantidade)': [],
                          'EXAMES (valor)': [],
                          'PRONTO-SOCORRO (quantidade)': [],
                          'PRONTO-SOCORRO (valor)': [],
                          'Desc por Co-Part. (-)': [],
                          'Bonificação GRES(-)': [],
                          'Acerto (-)': [],
                          'IOF': []})
    
    else:
        dfl = get_info_from_leftside(page)
        dfr = get_info_from_rightside(page)
        dfs = DataFrame({'Subestipulante': [vars.sub for _ in range(len(dfl.index))]})
        return pd.concat([dfs, dfl, dfr], axis = 'columns')
