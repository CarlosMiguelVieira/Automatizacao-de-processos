#%%
import os
import pandas as pd
import numpy as np
import re
from datetime import date
from src.ferramentas import mandar_email

# %%
dfFIMRTC = pd.read_excel(r'Q:\CM_teste\FIM.xls', index_col = False, sheet_name="Margem" )
print(dfFIMRTC)


print('oi')

LiqEligible = dfFIMRTC.Instrumento.str.contains('LiqEligible', flags = re.IGNORECASE, regex = True, na = False)
dfEliFIM=dfFIMRTC[LiqEligible]

LiqIneligible = dfFIMRTC.Instrumento.str.contains('LiqIneligible', flags = re.IGNORECASE, regex = True, na = False)
dfIneFIM=dfFIMRTC[LiqIneligible]

LiqCollat = dfFIMRTC.Instrumento.str.contains('LiqCollat', flags = re.IGNORECASE, regex = True, na = False)
dfCollatFIM=dfFIMRTC[LiqCollat]
print(dfCollatFIM)

EliFIM = dfEliFIM['Total'].sum()
IneFIM = dfIneFIM['Total'].sum()
CollatFIM = dfCollatFIM['Total'].sum()



# os.chdir(r'Q:\Total Return Fim')
# list_dir = os.listdir()
# fname = list_dir[-1]
# dfFIMPatrimonio = pd.read_excel(fname, index_col = False )
path = r'Q:\Total Return Fim'
# path = r'Q:\Carteira Diaria - 2022-06'

os.chdir(path)
cwd_str = os.getcwd()
list_dir = os.listdir()

list_rel = list(filter(lambda rls: 'ATIVA FIM EXCL_CARTEIRA_DIARIA_' in rls, list_dir))

fname = list_rel[-1]
# fname = 'ATIVA FIA_CARTEIRA_DIARIA_29062022.xlsx'
# pd.DataFrame(df.iloc[5:,:].values)

dfFIMPatrimonio = pd.read_excel(os.path.join(cwd_str, fname), index_col = False)

dfFIMPatrimonio = dfFIMPatrimonio[dfFIMPatrimonio['CARTEIRA DIARIA'].str.match('PATRIMONIO', na = False)]
dfFIMPatrimonio = dfFIMPatrimonio[['CARTEIRA DIARIA','Unnamed: 19']]

data_atual = str(date.today())
EliFIM
if EliFIM>0:
       EliFIM=0
IneFIM
if IneFIM>0:
       IneFIM=0
riscoTotalFim = EliFIM + IneFIM
CollatFIM
defSupFIM =riscoTotalFim+CollatFIM
patrimonioFIM = dfFIMPatrimonio['Unnamed: 19'].tolist()
perPLFIM = riscoTotalFim/float(patrimonioFIM[0])





dfFIM = pd.read_excel(r'Q:\Controle Risco FIM Total Return.xlsx', index_col = False )
# dfFIM.drop(columns=['Coluna1'], inplace=True)
dfFIM['Data'] = dfFIM['Data'].astype('str')
dfFIM.info()
dfFIM.loc[dfFIM.shape[0]]=[ data_atual,EliFIM,IneFIM,riscoTotalFim,CollatFIM,defSupFIM,float(patrimonioFIM[0]),perPLFIM ]




dfFIM['Data'] = pd.to_datetime(dfFIM['Data'], errors='coerce')
dfFIM['Data'] = dfFIM['Data'].dt.strftime('%d/%m/%Y')


dfFIM.to_excel(r'Q:\Controle Risco FIM Total Return.xlsx', index=False)


print(dfFIMPatrimonio)
dfFIM['%PL']=abs(dfFIM['%PL'])
dfFIM['%PL'] = pd.Series(["{0:.2f}%".format(val * 100) for val in dfFIM['%PL']], index = dfFIM.index)
dfFIM['PL']=dfFIM['PL'].astype('float64')
dfFIM['PL']=dfFIM['PL'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfFIM['Risco Total']=dfFIM['Risco Total'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfFIM['VaR 10D BVSP']=dfFIM['VaR 10D BVSP'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfFIM['VaR 10D BMF']=dfFIM['VaR 10D BMF'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfFIM['Garantias \nDepositadas']=dfFIM['Garantias \nDepositadas'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfFIM['Deficit / \nSuperavit']=dfFIM['Deficit / \nSuperavit'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfFIM=dfFIM.round(2)
##################################################


dfHEDGERTC = pd.read_excel(r'Q:\CM_teste\HEADGE.xls', index_col = False, sheet_name="Margem" )
print(dfHEDGERTC)



LiqEligible = dfHEDGERTC.Instrumento.str.contains('LiqEligible', flags = re.IGNORECASE, regex = True, na = False)
dfEliHEDGE=dfHEDGERTC[LiqEligible]

LiqIneligible = dfHEDGERTC.Instrumento.str.contains('LiqIneligible', flags = re.IGNORECASE, regex = True, na = False)
dfIneHEADGE=dfHEDGERTC[LiqIneligible]

LiqCollat = dfHEDGERTC.Instrumento.str.contains('LiqCollat', flags = re.IGNORECASE, regex = True, na = False)
dfCollatHEADGE=dfHEDGERTC[LiqCollat]
print(dfCollatHEADGE)

EliHEADGE = dfEliHEDGE['Total'].sum()
IneHEADGE = dfIneHEADGE['Total'].sum()
CollateHEADGE = dfCollatHEADGE['Total'].sum()


##################################################
# os.chdir(r'Q:\CM_teste\HEADGE')
# list_dir = os.listdir()
# fname = list_dir[-1]
# dfHEADGEPatrimonio = pd.read_excel(fname, index_col = False )
path = r'Q:\HEDGE FIM'
# path = r'Q:\Carteira Diaria - 2022-06'

os.chdir(path)
cwd_str = os.getcwd()
list_dir = os.listdir()

list_rel = list(filter(lambda rls: 'ATIVA FIM_CARTEIRA_DIARIA_' in rls, list_dir))

fname = list_rel[-1]
# fname = 'ATIVA FIA_CARTEIRA_DIARIA_29062022.xlsx'
# pd.DataFrame(df.iloc[5:,:].values)

dfHEADGEPatrimonio = pd.read_excel(os.path.join(cwd_str, fname), index_col = False)

dfHEADGEPatrimonio = dfHEADGEPatrimonio[dfHEADGEPatrimonio['CARTEIRA DIARIA'].str.match('PATRIMONIO', na = False)]
dfHEADGEPatrimonio = dfHEADGEPatrimonio[['CARTEIRA DIARIA','Unnamed: 19']]

data_atual = date.today()
EliHEADGE
if EliHEADGE>0:
       EliHEADGE=0
IneHEADGE
if IneHEADGE>0:
       IneHEADGE=0
riscoTotalHEADGE = EliHEADGE + IneHEADGE
CollateHEADGE
defSupHEADGE =riscoTotalHEADGE+CollateHEADGE
patrimonioHEADGE = dfHEADGEPatrimonio['Unnamed: 19'].tolist()
perPLHEADGE = riscoTotalHEADGE/float(patrimonioHEADGE[0])






dfHEDGE = pd.read_excel(r'Q:\Controle Risco HEDGE FIM.xlsx', index_col = False )
# dfHEDGE.drop(columns=['Coluna1'], inplace=True)
dfHEDGE['Data'] = dfHEDGE['Data'].astype('str')
dfHEDGE.loc[dfHEDGE.shape[0]]=[ data_atual,EliHEADGE,IneHEADGE,riscoTotalHEADGE,CollateHEADGE,defSupHEADGE,float(patrimonioHEADGE[0]),perPLHEADGE ]


dfHEDGE['%PL'] = dfHEDGE['%PL'] 

dfHEDGE['Data'] = pd.to_datetime(dfHEDGE['Data'], errors='coerce')
dfHEDGE['Data'] = dfHEDGE['Data'].dt.strftime('%d/%m/%Y')

# dfHEDGE['Risco Total'] = dfHEDGE['Risco Total'].replace('.',',')


dfHEDGE.info()
dfHEDGE.to_excel(r'Q:\Controle Risco HEDGE FIM.xlsx', index=False)
print(dfHEADGEPatrimonio)
dfHEDGE['%PL'] = abs(dfHEDGE['%PL'])
dfHEDGE['%PL'] = pd.Series(["{0:.2f}%".format(val * 100) for val in dfHEDGE['%PL']], index = dfHEDGE.index)
dfHEDGE['PL']=dfHEDGE['PL'].astype('float64')
dfHEDGE['PL']=dfHEDGE['PL'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfHEDGE['Risco Total']=dfHEDGE['Risco Total'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfHEDGE['VaR 10D BVSP']=dfHEDGE['VaR 10D BVSP'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfHEDGE['VaR 10D BMF']=dfHEDGE['VaR 10D BMF'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfHEDGE['Garantias \nDepositadas']=dfHEDGE['Garantias \nDepositadas'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
dfHEDGE['Deficit / \nSuperavit']=dfHEDGE['Deficit / \nSuperavit'].map('{:,.2f}'.format).str.replace(",", "~").str.replace(".", ",").str.replace("~", ".")
data = data_atual.strftime('%d/%m/%Y')
dfHEDGE=dfHEDGE.round(2)
##################################################
# %%
data_email=date.today()
data_email = data_email.strftime('%d/%m/%Y')
endereco = ""
endereco_cc = "cm_vp@hotmail.com"
titulo = "Controle Histórico - " + data_email
mensagem = ("<p>Prezados,</p>"
       "<p></p>"
       "<p>Segue abaixo a tabela do controle histórico de risco.</p>"
       "<p></p>"
       "<p>"+dfHEDGE.tail(5).to_html(index=False).replace('<td>', '<td align="center">')+"</p>")
mandar_email(endereco,endereco_cc,titulo,mensagem)






endereco = ""
endereco_cc = "cm_vp@hotmail.com"
titulo = "Controle Histórico de Risco - " + data_email
mensagem = ("<p>Prezados,</p>"
       "<p></p>"
       "<p>Segue abaixo a tabela do controle histórico de risco</p>"
       "<p>"+dfFIM.tail(5).to_html(index=False).replace('<td>', '<td align="center">')+"</p>")
# anexo = r'Q:\VaR_FIA.png'
mandar_email(endereco,endereco_cc,titulo,mensagem)
# %%
