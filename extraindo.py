import re
import pandas as pd

with open('SWCHD53_09-18-2024_17184.txt', 'r') as arquivo:
    conteudo = arquivo.read()

PANS = re.findall(r'PAN\s+([0-9]{19})', conteudo)# NUMERO DO CARTÃO/CONTA UTILIZADA NO PAGAMENTO
ADV_REASON = re.findall(r'ADV REASON\s+(.{8})\s+SWCH DATE', conteudo)# RAZÃO DA APRESENTAÇÃO / CHARGE BACK
ORG_DATE = re.findall(r'ORG DATE\s+([0-9]{2}-[0-9]{2})', conteudo)# DATA DA COMPRA
ORG_AMT = re.findall(r'ORG AMT\s+([0-9]{1,2}\.[0-9]{2})', conteudo)# VALOR DE COMPRA
ACQ_INST_NAMES = re.findall(r'ACQ INST NAME\s+(.{30})', conteudo)# NOME DO ESTABELECIMENTO

#RETIRANDO OS ESPAÇOS INDESEJADOS DOS NOMES DAS INSTITUIÇÕES
for i in range(len(ACQ_INST_NAMES)):
    ACQ_INST_NAMES[i] = ACQ_INST_NAMES[i].rstrip()

#CRIANDO UM DICIONARIO PARA A CRIAÇÃO DO DATAFRAME
infos = {
    'PANS':PANS,
    'ADV_REASON':ADV_REASON,
    'ORG_DATE':ORG_DATE,
    'ORG_AMT':ORG_AMT,
    'ACQ_INST_NAMES':ACQ_INST_NAMES
}

    
df = pd.DataFrame(infos)

df.to_excel('planilha.xlsx', index=False, header=False)
