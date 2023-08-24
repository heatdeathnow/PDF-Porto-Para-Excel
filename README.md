# PDF-Porto-Para-Excel
Programa que pega um PDF de faturamento da Porto Seguro e o formatada em `.xlsx` enquanto verifica a exatidão dos valores.

### Funcionamento
O programa recorta o PDF em várias fatias para lê-lo sem interferências. Ele extrai a informação dos beneficiáios das páginas, junto com qual o subestipulante, e então concatena num Dataframe principal. Quando ele chega no final do PDF, ele escreve toda a informação do Dataframe numa planilha `.xlsx` com o mesmo nome do arquivo `.pdf` original.

São então adicionadas outras quatro planilhas com cálculos que, em tese, deveriam bater com os valores cobrados no faturamento da Porto. Formatação condicional é adicionada de tal forma que, se os valores divergirem a célula fica vermelho, senão fica verde.

O arquivo é, então, formatado com grades, cores, e formatação numérica onde consta.
