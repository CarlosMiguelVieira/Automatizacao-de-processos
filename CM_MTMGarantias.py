#%%
from ast import Index
from operator import index
import os
import string
import pandas as pd
import zipfile
import numpy as np
from src.ferramentas import mandar_email
from datetime import date
#%%%
# path = r'C:\Users\cpinto\Desktop\MTMs'
# # path = r'Q:\Risco\Novo Risco\Ativa Asset\FIA\Carteira Diaria - 2022-06'

# os.chdir(path)

# list_dir = os.listdir()
# fname = list_dir[-1]
# print(fname)

      
# fantasy_zip = zipfile.ZipFile('C:\\Users\\cpinto\\Desktop\\MTMs\\' + fname)
# fantasy_zip.extract('MaximumTheoreticalMargin.csv', 'C:\\Users\\cpinto\\Desktop\\MTMs')
# fantasy_zip.close()


# list=['WINQ22','WDOU22','DOLU22','DOLQ22','INDQ22','WSPU22','CCMU22','BGIU22','SJCU22','ICFU22']
list_ticker = []
list_garantia =[]
list_disagio = []
list_ok = []
dflista = pd.read_csv(r'Q:\Risco\Novo Risco\pythonrisco\Codigos\data\CM_MTM\IBXX.csv',index_col = False,sep=';', encoding='latin-1',names= [ 'Codigo',1,2,3,4,5])
df = pd.read_csv(r'Q:\Risco\Novo Risco\pythonrisco\Codigos\data\MTM_export\MaximumTheoreticalMargin.csv', sep=';',index_col = False, names= [ 'MTM', 'Data', 'Minima', 'Maxima', 'Minima2', 'Ticker',"AUX"])


dflista = dflista['Codigo'].loc[2:99].tolist()
list = dflista


for ticker in list:
   
   
    #### CRIANDO UMA COLUNA AUXILIAR
    df['AUX'] = range(df.shape[0])
    df['AUX']

    #### ACHANDO A LINHA DE REFERENCIA DO TICKER
    df_selecao = df.loc[df['Ticker'].str.endswith(ticker, na=False)] #print (df.loc[df['Ticker'] == 'PETR4']) OUTRA FORMA DE LER
   

    #### CRIANDO VARIAVEL PARA ENCONTRAR O VALOR DA GARANTIA
    x = df_selecao['AUX'].tolist()

    #### CRIANDO AS VARIAVEIS DO VALOR DE GARANTIA E PREÇO DE FECHAMENTO
    minima=df.loc[x[0]+2:x[0]+2,'Minima'].tolist()
    minima2=df.loc[x[0]+2:x[0]+2,'Minima2'].tolist()

    maxima=df.loc[x[0]+2:x[0]+2,'Maxima'].tolist()
    
    preco=df_selecao['Minima2'].tolist()
    
    #### TRANSFORMANDO OS DADOS DA GARANTIA DE STRING PARA FLOAT

    # isNaN = pd.isna(minima[0])

    # if (isNaN == True or minima[0] == '0'):
    #     minimaFloat=1
    #     maximaFloat=1
    #     precoFloat=1
    # else:
    try:
        minimaFloat=float(minima[0].replace(',','.'))
        maximaFloat=float(maxima[0].replace(',','.'))
        precoFloat=float(preco[0].replace(',','.'))
    except:
         minimaFloat=float(minima2[0].replace(',','.'))
         maximaFloat=float(minima2[0].replace(',','.'))
         precoFloat=float(preco[0].replace(',','.'))
    # except
    # print(minimaFloat,maximaFloat,precoFloat)
    # print(minima[0])

    # isNaN = pd.isna(minima[0])
    # print(isNaN)


    #### Pegar o menor valor
    if minimaFloat>=maximaFloat:
        garantia = maximaFloat
    else:
        garantia = minimaFloat


    # print(garantia)
    # print(precoFloat)
    disagio=abs((garantia/precoFloat)-1)
    # print(ticker,garantia,disagio)


    list_garantia.append(garantia)
    list_ticker.append(ticker)
    list_disagio.append(disagio)
    





dfListaGarantia = pd.read_excel(r'Q:\Risco\Novo Risco\pythonrisco\Codigos\data\CM_MTM\Aceitagarantia.xlsx',index_col = False)
dfListaGarantia = dfListaGarantia['Código'].tolist()

for x in list_ticker:

    for i in dfListaGarantia:
        if x == i:
            ok = True
            break
        else:
            ok = False
    list_ok.append(ok)


dfl = pd.DataFrame((zip(list_ticker,list_garantia,list_disagio,list_ok)), columns = ['Ticker','Garantia','Desagio','Aceito como garantia?'])
dfl['Aceito como garantia?']=dfl['Aceito como garantia?'].replace(True,'Sim').replace(False,'Não')

dfl.round(2).to_excel(r'Q:\Risco\Novo Risco\pythonrisco\Codigos\data\CM_MTM\Valor Minimo em Garantia.xlsx', sheet_name='MTM', index=False)
# %%
today=date.today()
data = today.strftime('%d/%m/%Y')
endereco = "options@ativainvestimentos.com.br;mesainstrj@ativainvestimentos.com.br;comercialbh@ativainvestimentos.com.br>;comercialrj@ativainvestimentos.com.br;comercialrs@ativainvestimentos.com.br;comercialsp@ativainvestimentos.com.br;comercialpr@ativainvestimentos.com.br;comerciais@ativainvestimentos.com.br;atendimentorj@ativainvestimentos.com.br;custodia@ativainvestimentos.com.br;comercialba@ativainvestimentos.com.br"
endereco_cc = "risco@ativainvestimentos.com.br"
titulo = "Valor mínimo em garantia - " + str(data)
mensagem = ("<p>Prezados,</p>"
       "<p>      Segue em anexo a análise da carteira do IBX-100 em relação valor mínimo em garantia por ação. </p>"
       "<p>Contendo na planilha os seguintes parâmetros:</p>"
       "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp• Garantia = Valor do papel alocado na carteira de garantia</p>"
       "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp• Deságio = Diferença entre o valor real e o nominal do ativo  </p>"
       "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp• Aceito em garantia = Possibilidade do ativo ser alocado em garantia</p>"
    )
anexo = r'Q:\Risco\Novo Risco\pythonrisco\Codigos\data\CM_MTM\Valor Minimo em Garantia.xlsx'
mandar_email(endereco,endereco_cc,titulo,mensagem,anexo)


# %%
