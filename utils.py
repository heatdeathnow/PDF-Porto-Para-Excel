from pdfplumber.page import Page
from pandas import DataFrame
import vars


def update_sub(page: Page):
    cropped = page.within_bbox((0, 70, 375, 80))
    vars.sub = cropped.extract_text().replace('Subestipulante:', '').strip()

def fix_types(df: DataFrame) -> DataFrame:
    for i in range(len(df.index)):
        if df.loc[i, 'Idade'] != '':
            df.loc[i, 'Idade'] = int(df.loc[i, 'Idade'])

        if df.loc[i, 'Prêmio Base'] != '':
            df.loc[i, 'Prêmio Base'] = float(df.loc[i, 'Prêmio Base'].replace('.', '').replace(',', '.'))

        if df.loc[i, 'CONSULTAS (quantidade)'] != '':
            df.loc[i, 'CONSULTAS (quantidade)'] = int(df.loc[i, 'CONSULTAS (quantidade)'])
        
        if df.loc[i, 'CONSULTAS (valor)'] != '':
            df.loc[i, 'CONSULTAS (valor)'] = float(df.loc[i, 'CONSULTAS (valor)'].replace('.', '').replace(',', '.'))
        
        if df.loc[i, 'EXAMES (quantidade)'] != '':
            df.loc[i, 'EXAMES (quantidade)'] = int(df.loc[i, 'EXAMES (quantidade)'])

        if df.loc[i, 'EXAMES (valor)'] != '':
            df.loc[i, 'EXAMES (valor)'] = float(df.loc[i, 'EXAMES (valor)'].replace('.', '').replace(',', '.'))

        if df.loc[i, 'PRONTO-SOCORRO (quantidade)'] != '':
            df.loc[i, 'PRONTO-SOCORRO (quantidade)'] = int(df.loc[i, 'PRONTO-SOCORRO (quantidade)'])

        if df.loc[i, 'PRONTO-SOCORRO (valor)'] != '':
            df.loc[i, 'PRONTO-SOCORRO (valor)'] = float(df.loc[i, 'PRONTO-SOCORRO (valor)'].replace('.', '').replace(',', '.'))

        if df.loc[i, 'Pro-Rata (quantidade)'] != '':
            df.loc[i, 'Pro-Rata (quantidade)'] = int(df.loc[i, 'Pro-Rata (quantidade)'])
        
        if df.loc[i, 'Pro-Rata (valor)'] != '':
            df.loc[i, 'Pro-Rata (valor)'] = float(df.loc[i, 'Pro-Rata (valor)'].replace('.', '').replace(',', '.'))
        
        if df.loc[i, 'Desc por Co-Part. (-)'] != '':
            df.loc[i, 'Desc por Co-Part. (-)'] = float(df.loc[i, 'Desc por Co-Part. (-)'].replace('.', '').replace(',', '.'))
        
        if df.loc[i, 'Bonificação GRES(-)'] != '':
            df.loc[i, 'Bonificação GRES(-)'] = float(df.loc[i, 'Bonificação GRES(-)'].replace('.', '').replace(',', '.'))
        
        if df.loc[i, 'Acerto (-)'] != '':
            df.loc[i, 'Acerto (-)'] = float(df.loc[i, 'Acerto (-)'].replace('.', '').replace(',', '.'))
        
        if df.loc[i, 'IOF'] != '':
            df.loc[i, 'IOF'] = float(df.loc[i, 'IOF'].replace('.', '').replace(',', '.'))
