from pdfplumber.page import Page
from pandas import DataFrame
import vars


def to_num(txt: str) -> float | int:
    if ',' in txt:
        return float(txt.replace('.', '').replace(',', '.').strip())
    
    else:
        return int(txt.replace('.', '').strip())

def get_sub(page: Page) -> str:
    cropped = page.within_bbox((0, 70, 375, 80))
    return cropped.extract_text().replace('Subestipulante:', '').strip()

def fix_types(df: DataFrame) -> DataFrame:
    for i in range(len(df.index)):
        if df.loc[i, 'Idade'] != '':
            df.loc[i, 'Idade'] = int(df.loc[i, 'Idade'])

        if df.loc[i, 'Prêmio Base'] != '':
            df.loc[i, 'Prêmio Base'] = to_num(df.loc[i, 'Prêmio Base'])

        if df.loc[i, 'CONSULTAS (quantidade)'] != '':
            df.loc[i, 'CONSULTAS (quantidade)'] = int(df.loc[i, 'CONSULTAS (quantidade)'])
        
        if df.loc[i, 'CONSULTAS (valor)'] != '':
            df.loc[i, 'CONSULTAS (valor)'] = to_num(df.loc[i, 'CONSULTAS (valor)'])
        
        if df.loc[i, 'EXAMES (quantidade)'] != '':
            df.loc[i, 'EXAMES (quantidade)'] = int(df.loc[i, 'EXAMES (quantidade)'])

        if df.loc[i, 'EXAMES (valor)'] != '':
            df.loc[i, 'EXAMES (valor)'] = to_num(df.loc[i, 'EXAMES (valor)'])

        if df.loc[i, 'PRONTO-SOCORRO (quantidade)'] != '':
            df.loc[i, 'PRONTO-SOCORRO (quantidade)'] = int(df.loc[i, 'PRONTO-SOCORRO (quantidade)'])

        if df.loc[i, 'PRONTO-SOCORRO (valor)'] != '':
            df.loc[i, 'PRONTO-SOCORRO (valor)'] = to_num(df.loc[i, 'PRONTO-SOCORRO (valor)'])

        if df.loc[i, 'Pro-Rata (quantidade)'] != '':
            df.loc[i, 'Pro-Rata (quantidade)'] = int(df.loc[i, 'Pro-Rata (quantidade)'])
        
        if df.loc[i, 'Pro-Rata (valor)'] != '':
            df.loc[i, 'Pro-Rata (valor)'] = to_num(df.loc[i, 'Pro-Rata (valor)'])
        
        if df.loc[i, 'Desc por Co-Part. (-)'] != '':
            df.loc[i, 'Desc por Co-Part. (-)'] = to_num(df.loc[i, 'Desc por Co-Part. (-)'])
        
        if df.loc[i, 'Bonificação GRES(-)'] != '':
            df.loc[i, 'Bonificação GRES(-)'] = to_num(df.loc[i, 'Bonificação GRES(-)'])
        
        if df.loc[i, 'Acerto (-)'] != '':
            df.loc[i, 'Acerto (-)'] = to_num(df.loc[i, 'Acerto (-)'])
        
        if df.loc[i, 'IOF'] != '':
            df.loc[i, 'IOF'] = to_num(df.loc[i, 'IOF'])
