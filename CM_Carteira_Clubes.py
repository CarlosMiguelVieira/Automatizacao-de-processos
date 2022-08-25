# %%
import pandas as pd
import numpy as np
# %%
dflista = pd.read_csv(r'Q:\Risco x Backoffice\Clubes\Bases\lista.csv',index_col = False,sep=';')
dflista = dflista['clube'].tolist()

list_ok=[]
list_clube=[]
carteira = pd.DataFrame(columns=['Código','QTDE_CLUBE','QTDE_PORTAL','Diferenca'])


for clube in dflista:
    clube = str(clube)
    # clube = '106824'
    # clube = '96144'
    df = pd.read_excel(rf'Q:\Risco x Backoffice\Clubes\Arquivos Clubes\{clube}.xlsx', index_col = False )
    # df = df[['Unnamed: 1','Unnamed: 31']].dropna()

    # df = pd.read_excel(r'C:\Users\cpinto\Desktop\Teste\Carteira_106824_11-8-2022.xlsx', index_col = False )
    # # df = df2[['Unnamed: 1','Unnamed: 37']].dropna()

    # df2 = pd.read_excel(r'Q:\Risco\Novo Risco\1 - Rotinas\Clubes\Carteira_419572.xlsx', index_col = False )
    # # df = df[['Unnamed: 1','Unnamed: 22']].dropna()

    df.drop(df.index[0:11], inplace = True)  
    df_header = df.iloc[0] #grab the first row for the header
    df = df[1:] #take the data less the header row
    df.columns = df_header #set the header row as the df header
    df.drop(df.index[0], inplace = True)  
    df = df[['Código','Quantidade']].dropna()
    df.drop(df.index[len(df.index)-1], inplace= True)
    df['Quantidade'] = df['Quantidade'].astype(int)


    dfBase = pd.read_csv(r'Q:\Risco x Backoffice\Clubes\Bases\Base_SQL.csv',sep=';')
    dfBase=dfBase.groupby(['COD_CLI','COD_NEG']).agg({'QTDE_TOT':'sum'}).reset_index()

    dfBase['COD_CLI']=dfBase['COD_CLI'].astype(str)
    dfBaseClub=dfBase[dfBase['COD_CLI'].str.contains(clube)] 
    dfBaseClub=dfBaseClub[['COD_NEG','QTDE_TOT']]


    dfBaseClub.reset_index(drop=True, inplace=True)
    df.reset_index(drop=True, inplace=True)
    len(df)
    
    df.rename(columns = {'Código':'COD_NEG', 'Quantidade':'QTDE_TOT'}, inplace = True)
    df.sort_values(by='COD_NEG',inplace=True)
    dfBaseClub.sort_values(by='COD_NEG',inplace=True)
    list_certo=[]
    if len(df) == len(dfBaseClub):
        for i in list(range(len(df))):
            
            if (dfBaseClub.loc[0+i:0+i,'COD_NEG'] == df.loc[0+i:0+i,'COD_NEG']).bool() & (dfBaseClub.loc[0+i:0+i,'QTDE_TOT'] == df.loc[0+i:0+i,'QTDE_TOT']).bool():
                list_certo.append(True)
                         


        if len(list_certo) == len(df):
            list_ok.append('OK')
            list_clube.append(clube)
            print("if2")
        else:
            list_ok.append('Tem ativo diferente ou quantidade diferente')
            list_clube.append(clube)
            print("else1")
            dfDiferente = pd.merge(dfBaseClub, df, how = 'outer', on = 'COD_NEG')            
            dfDiferente['Diferenca'] = dfDiferente['QTDE_TOT_x'] - dfDiferente['QTDE_TOT_y']
            diferenca =   dfDiferente['Diferenca'] != 0
            dfDiferente=dfDiferente[diferenca].fillna('Não tem')
            dfDiferente.rename(columns = {'COD_NEG':'Código', 'QTDE_TOT_x':'QTDE_PORTAL', 'QTDE_TOT_y':'QTDE_CLUBE'}, inplace = True)
            lista_else = []
            for y in list(range(len(dfDiferente))):
                lista_else.append(clube)
            dfDiferente['Clube'] = lista_else
            
            carteira = pd.concat([carteira,dfDiferente])
            


    else:

        if len(df) >= len(dfBaseClub):
            list_ok.append(f'Tem {len(df) - len(dfBaseClub)} ativos a mais no portal')
            list_clube.append(clube) 
            dfMaisPortal = pd.merge(dfBaseClub, df, how = 'right', on = 'COD_NEG')
            dfMaisPortal['Diferenca'] = dfMaisPortal['QTDE_TOT_x'] - dfMaisPortal['QTDE_TOT_y']
            diferenca =   dfMaisPortal['Diferenca'] != 0
            dfMaisPortal=dfMaisPortal[diferenca].fillna('Não tem')
            dfMaisPortal.rename(columns = {'COD_NEG':'Código', 'QTDE_TOT_x':'QTDE_PORTAL', 'QTDE_TOT_y':'QTDE_CLUBE'}, inplace = True)
            lista_if1 = []
            for y in list(range(len(dfMaisPortal))):
                lista_if1.append(clube)
            dfMaisPortal['Clube'] = lista_if1
            #Merge para comparar quem não tem no outro/ outer
            carteira = pd.concat([carteira,dfMaisPortal])
            
        else:        
            list_ok.append(f'Tem {len(dfBaseClub) - len(df)} ativos a mais no clube')
            list_clube.append(clube)   
            dfMaisClub = pd.merge(dfBaseClub, df, how = 'left', on = 'COD_NEG')
            dfMaisClub['Diferenca'] = dfMaisClub['QTDE_TOT_x'] - dfMaisClub['QTDE_TOT_y']
            diferenca =   dfMaisClub['Diferenca'] != 0
            dfMaisClub=dfMaisClub[diferenca].fillna('Não tem')
            dfMaisClub.rename(columns = {'COD_NEG':'Código', 'QTDE_TOT_x':'QTDE_PORTAL', 'QTDE_TOT_y':'QTDE_CLUBE'}, inplace = True)
            lista_else2 = []
            for y in list(range(len(dfMaisClub))):
                lista_else2.append(clube)
            dfMaisClub['Clube'] = lista_else2

            
            carteira = pd.concat([carteira,dfMaisClub])
print(list_clube,list_ok)




dfRelatorio = carteira[['Clube','Código','QTDE_CLUBE','QTDE_PORTAL','Diferenca']].reset_index(drop=True)
dfRelatorio.to_excel(r'Q:\Risco x Backoffice\Clubes\Relatorio Descritivo.xlsx', index=False)



dfl = pd.DataFrame((zip(list_clube,list_ok)), columns = ['Clube','Situação'])

dfl.to_excel(r'Q:\Risco x Backoffice\Clubes\Relatorio Simplificado.xlsx', index=False)
# %%