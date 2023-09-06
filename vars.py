from os import cpu_count


max_threads = min(32, cpu_count() + 4)
main_threads = 4

nformat = 'R$* #,##0.00;-R$* #,##0.00;-;'
sub_has_changed = False
sub = ''

qev = ''
premio = ''
total = ''
detalhe = ''

# Para caso mude
SUB = 'A'  # Subestipulante
SEG = 'B'  # Seguro
DEP = 'C'  # Departamento
NOM = 'D'  # Nome do segurado
REG = 'E'  # Reg. Func.
IDA = 'F'  # Idade
PAR = 'G'  # Parentesco
PLA = 'H'  # Plano
MOV = 'I'  # Mov./Inic.Vig.
SIT = 'J'  # Situação do Segurado
PRE = 'K'  # Prêmio Base
RAQ = 'L'  # Pro-Rata (quantidade)
RAV = 'M'  # Pro-Rata (valor)
COQ = 'N'  # CONSULTAS (quantidade)
COV = 'O'  # CONSULTAS (valor)
EXQ = 'P'  # EXAMES (quantidade)
EXV = 'Q'  # EXAMES (valor)
PRQ = 'R'  # PRONTO-SOCORRO (quantidade)
PRV = 'S'  # PRONTO-SOCORRO (valor)
DES = 'T'  # Desc por Co-Part. (-)
BON = 'U'  # Bonificação GRES(-)
ACE = 'V'  # Acerto (-)
IOF = 'W'  # IOF
COP = 'X'  # Coparticipação
TDE = 'Y'  # TOTAL DO DEP.
PRR = 'Z'  # Prêmio Real
