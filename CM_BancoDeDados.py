
import pandas as pd
import numpy as np
import cx_Oracle
import os
from datetime import date
from src.ferramentas import mandar_email
import win32com.client  
import datetime
data = date.today()

conn = cx_Oracle.connect(user="usuario", password="senha",dsn="local:1521/dnsqualquer")
cur = conn.cursor()

### SFP
sql = "select cd_cpfcgc,sum(vl_ben) from corrwin.tscsfp group by cd_cpfcgc"
result = cur.execute(sql).fetchall()
dfSFP = pd.DataFrame(result,columns=['CD_CPFCGC','SUM(VL_BEN)'])
dataSFP = data.strftime('%d.%m.%Y')


###FLEX
sql = "select codcli,NATOPE, case when COMMOD = 'FPA' THEN 'PUT' when COMMOD = 'FCA' THEN 'CALL' ELSE COMMOD END COMMOD, DATREG, DATVCT, TAMBAS, PREEXE, VLRNEG, VARIAVEL from CORRWIN.TMFPOSIC_FLX cad where dt_datmov = trunc(sysdate-1)"
result = cur.execute(sql).fetchall()
dfFlex = pd.DataFrame(result,columns=['CODCLI','NATOPE','COMMOD', 'DATREG', 'DATVCT', 'TAMBAS', 'PREEXE', 'VLRNEG', 'VARIAVEL'])

# x = datetime.datetime.now()
ontem = datetime.datetime.now()- datetime.timedelta(days=1)
x=ontem.strftime("%d/%b/%Y")
sql = f"SELECT pat.DT_REFERENCIA, SUM(SALDODISPONIVEL) AS Saldo FROM portalativa.risco_patrimonio_cliente_h pat WHERE CD_CPFCGC <> '33775974000104' AND pat.DT_REFERENCIA = '{x}' AND SALDODISPONIVEL > 0 GROUP BY pat.DT_REFERENCIA"
result = cur.execute(sql).fetchall()
dfSaldo = pd.DataFrame(result,columns=['pat.DT_REFERENCIA','SUM(SALDODISPONIVEL)'])
dataSaldo = data.strftime('%d-%m-%Y')


dfSaldo.to_excel(fr'Q:\Saldo de terceiros\Saldo_terceiros {dataSaldo}.xlsx', index=False)
dfFlex.to_excel(fr'Q:\Flex\Flex - {data}.xls', index=False)
dfSFP.to_csv(fr'Q:\PLDFT\SFP\SFP - {data}.csv', index=False)






data = data.strftime('%d/%m/%Y')
endereco = ""
endereco_cc = "cm_vp@hotmail.com"
titulo = "Cotrole rotina da manhã - " + str(data)
mensagem = ("<p>Prezados,</p>"
       "<p></p>"
       "<p>Arquivos do Banco de dados enviados com sucesso.</p>"
)

olMailItem = 0x0
obj = win32com.client.Dispatch('Outlook.Application')
newMail = obj.CreateItem(olMailItem)
newMail.Subject = titulo
newMail.HTMLBody = mensagem
newMail.To = endereco_cc
# newMail.display()
newMail.Send()

# anexo = r'Q:\VaR_FIA.png'
# mandar_email(endereco,endereco_cc,titulo,mensagem)


# data_email = data_email.strftime('%d/%m/%Y')
# endereco = ""
# endereco_cc = ""
# titulo = "Controle Histórico de Risco Ativa Asset HEDGE FIM - " + data_email
# mensagem = ("<p>Prezados,</p>"
#        "<p></p>"

# # anexo = r'Q:\VaR_FIA.png'
# mandar_email(endereco,endereco_cc,titulo,mensagem)






