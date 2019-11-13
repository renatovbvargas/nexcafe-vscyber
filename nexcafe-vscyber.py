import pandas as pd
import numpy as np
import fdb
import time
import os
import sys
import shutil
import zipfile
from datetime import datetime
from unicodedata import normalize

if os.path.exists(os.getcwd() + '\\VSCyber.FDB'):
    os.remove(os.getcwd() + '\\VSCyber.FDB')

shutil.copyfile(os.getcwd() + '\\_.FDB', os.getcwd() + '\\VSCyber.FDB')

file = os.getcwd() + '\\Exportar.xls'

hora = float(sys.argv[1].replace(',', '.'))

def timeToInt(strTime):

    if isinstance(strTime, int):
        return strTime / 60

    fmt = ''
    strTime = str(strTime)

    if strTime == '':
        return 0

    if strTime.find('h') != -1:
        fmt += "%Hh"

    if strTime.find('m') != -1:
        fmt += "%Mm"

    if strTime.find('s') != -1:
        fmt += "%Ss"

    if fmt == '':
        return 0

    try:
        result = time.strptime(strTime, fmt)
    except ValueError as e:
        result = time.strptime('0s', '%Ss')

    return result.tm_hour + (result.tm_min/60) + ((result.tm_sec/60)/60)


def insertPESSXFORMACNTT(cur, idformacntt, referencia, unidgeo, firstName, lastName, username):
    if(pd.notnull(referencia)):
        sql = "INSERT INTO PESSXFORMACNTT(Idpessxformacntt,Idpessoa,idformacntt,referencia,idlocd,Idinc,Dhinc,Idalt,Dhalt) select first 1 GEN_ID(PESSXFORMACNTT_GEN,1), p.IDPESSOA, {}, '{}', (select first 1 idunidgeo from unidgeo where nome='{}' and idunidgeo in (select idlocd from locd)) , 1, CURRENT_TIMESTAMP, NULL, NULL from pessoa p left join login l on p.idpessoa=l.idlogin where p.NOMEFANTASIA like '{}' and p.NOMECOMPLETO like '{}' and l.login like CASE WHEN '{}' = '' THEN p.idpessoa ELSE '{}' END".format(
            idformacntt, referencia, unidgeo, firstName, lastName, username, username)
        cur.execute(sql)


dfExportar = pd.read_excel(file, sheet_name=0, header=None)

dfExportar.columns = ['Nome', 'Username', 'Código', 'Status', 'Tipo', 'Débito', 'Cred.Tempo', 'Cred.Valor', 'Créditos Promocionais', 'Data Nasc.',
                      'Tempo Usado', 'RG', 'Endereço', 'Bairro', 'Cidade', 'UF', 'CEP', 'Sexo', 'E-mail', 'Telefone', 'Escola', 'NickName', 'Celular', 
                      'Incluído Em', 'Limite Débito', 'Incluído Por', 'Alterado Em', 'Alterado Por', 'Tit. Eleitor', 'Pai', 'P.Disponíveis', 'P. Acumulados',
                      'P. Resgatados', 'Mãe', 'Censura de Horário', 'CPF']

dfExportar = dfExportar[dfExportar.Username.notnull()]
dfExportar = dfExportar[dfExportar.Username.str.match('Username', na=False)==False]

dfExportar = dfExportar.replace("'","",regex=True)

new = dfExportar.Nome.str.split(" ", n=1, expand=True)
dfExportar['FirstName'] = new[0].str[:20]
dfExportar['LastName'] = new[1].str[:50]
dfExportar['LastName'] = dfExportar['LastName'].replace([None], [''])

dfExportar['Cred.Tempo'].fillna(0, inplace=True)
dfExportar['Cred.Valor'].fillna(0, inplace=True)
dfExportar['Débito'].fillna(0, inplace=True)

dfExportar['Valor'] = round((dfExportar['Cred.Tempo'].apply(timeToInt) * hora) +
                            dfExportar['Cred.Valor'] - dfExportar['Débito'], 2)

dfExportar['Valor'].fillna(0, inplace=True)

dfExportar['Créditos Promocionais'] = dfExportar['Créditos Promocionais'].replace([
                                                                                  None], [''])
dfExportar['Cortesia'] = round(dfExportar['Créditos Promocionais'].apply(timeToInt)* hora,2)

dfExportar['Data Nasc.'] = dfExportar['Data Nasc.'].replace([None], [''])
dfExportar['Data Nasc.'] = dfExportar['Data Nasc.'].replace([0], [''])

dfExportar['DataNasc'] = dfExportar['Data Nasc.'].apply(
    lambda x: None if str(x) == '' else datetime.strftime(x, '%Y.%m.%d'))

dfExportar.UF = dfExportar.UF.fillna('UF').str.upper()
dfExportar.Cidade = dfExportar.Cidade.fillna('Cidade')
dfExportar.Bairro = dfExportar.Bairro.fillna('Bairro')

dfExportar.Pai = dfExportar.Pai.replace([None], [''])
dfExportar.Mãe = dfExportar.Mãe.replace([None], [''])
dfExportar['Responsavel'] = np.where(
    dfExportar.Pai=='', dfExportar.Pai.str[:50], dfExportar.Mãe.str[:50])

dfExportar.Username = dfExportar.Username.astype(str).str.strip()

con = fdb.connect(dsn="localhost:{}\\VSCyber.FDB".format(os.getcwd()),
                  user="sysdba",
                  password="masterkey",
                  port=3050)

cur = con.cursor()

for index, row in dfExportar.iterrows():
    cur.execute("insert into pessoa values (GEN_ID(PESSOA_GEN,1), ?, ?, 'F', 0, 1, current_timestamp, 1, current_timestamp)",
                (row.FirstName, row.LastName))

    cur.execute("insert into pessoafisica (idpessoa, sexo) select idpessoa, ? from pessoa where nomefantasia = ? and Nomecompleto = ? AND IDPESSOA NOT IN ( SELECT Idpessoa FROM pessoafisica)",
                (row.Sexo, row.FirstName, row.LastName))

cur.execute("insert into cli (idcli, sitppgcli, flags) select idpessoa, 1, 2 from pessoa where idpessoa not in (select idcli from cli)")

cur.execute("update VRFXTABHORA set valor = {}".format(hora))

for UF in dfExportar.UF.unique():
    sql1 = "insert into UNIDGEO values (GEN_ID(UnidGeo_GEN,1),'{}', 1, current_timestamp, NULL, NULL)".format(
        UF)
    cur.execute(sql1)
    sql2 = "insert into UF (IDUF, sigla) select idunidgeo, '{}' from unidgeo where nome ='{}'".format(
        UF, UF)
    cur.execute(sql2)

for index, row in dfExportar[dfExportar.UF.notnull()].drop_duplicates(['UF', 'Cidade'])[['UF', 'Cidade']].iterrows():
    sql1 = "insert into UNIDGEO values (GEN_ID(UnidGeo_GEN,1),'{}', 1, current_timestamp, NULL, NULL)".format(
        row.Cidade)
    cur.execute(sql1)
    sql2 = "insert into LOCD (IDLOCD, IDUF) select first 1 idunidgeo, (SELECT first 1 idunidgeo FROM unidgeo WHERE nome = '{}' and idunidgeo in (select idUF from UF))  from unidgeo where nome ='{}' and idunidgeo not in (select iduf from UF) and idunidgeo not in (select idlocd from locd)".format(row.UF, row.Cidade)
    cur.execute(sql2)

for index, row in dfExportar[dfExportar.Cidade != ''].drop_duplicates(['Cidade', 'Bairro'])[['Cidade', 'Bairro']].iterrows():
    sql1 = "insert into UNIDGEO values (GEN_ID(UnidGeo_GEN,1),'{}', 1, current_timestamp, NULL, NULL)".format(
        row.Bairro)
    cur.execute(sql1)
    sql2 = "insert into BAIRRO (IDBAIRRO, IDLOCD) select first 1 idunidgeo, (SELECT first 1 idunidgeo FROM unidgeo WHERE nome = '{}' and idunidgeo in (select idLOCD from LOCD))  from unidgeo where nome ='{}' and idunidgeo not in (select iduf from UF) and idunidgeo not in (select idLOCD from LOCD) and idunidgeo not in (select idbairro from bairro)".format(row.Cidade, row.Bairro)
    cur.execute(sql2)

prevUsername = ''
dfExportar = dfExportar.sort_values('Username')
for index, row in dfExportar.iterrows():

    if(row.Tipo == 'Acesso Grátis'):
        cur.execute("update cli set BFree=1 where idcli in (select idpessoa from pessoa where NOMEFANTASIA like ? and NOMECOMPLETO like ?)",
                    (row.FirstName, row.LastName))

    if(row.Username.upper() == 'ADMIN'):
        row.Username += '_1'
    
    if(row.Username.strip() == prevUsername):
        row.Username += '*'

    prevUsername = row.Username

    cur.execute("INSERT INTO login (IdLogin,Login,PW,Flags) select first 1 p.idpessoa, ?, NULL, NULL from pessoa p left join mov m on p.idpessoa=m.idcli where p.NOMEFANTASIA like ? and p.NOMECOMPLETO like ? and p.idpessoa not in (select idlogin from login)",
                (row.Username, row.FirstName, row.LastName))

    if(row.Valor > 0):
        sql = "INSERT INTO Mov(IdMov,IdCli,DhMov,Valor,SiTpOpMov,IdCon,DtValidCred,IdInc,DhInc) select first 1 GEN_ID(Mov_GEN,1), p.IDPESSOA, CURRENT_TIMESTAMP, {}, 1, NULL, CURRENT_TIMESTAMP, 1, CURRENT_TIMESTAMP from pessoa p left join login l on p.idpessoa=l.idlogin where p.NOMEFANTASIA like '{}' and p.NOMECOMPLETO like '{}' and l.login like CASE WHEN '{}' = '' THEN p.idpessoa ELSE '{}' END and p.idpessoa NOT IN ( SELECT idcli FROM mov)".format(
            row.Valor, row.FirstName, row.LastName, row.Username, row.Username)
        cur.execute(sql)

    if(row.Cortesia > 0):
        sql = "INSERT INTO Mov(IdMov,IdCli,DhMov,Valor,SiTpOpMov,IdCon,DtValidCred,IdInc,DhInc) select first 1 GEN_ID(Mov_GEN,1), p.IDPESSOA, CURRENT_TIMESTAMP, {}, 6, NULL, CURRENT_TIMESTAMP, 1, CURRENT_TIMESTAMP from pessoa p left join login l on p.idpessoa=l.idlogin where p.NOMEFANTASIA like '{}' and p.NOMECOMPLETO like '{}' and l.login like CASE WHEN '{}' = '' THEN p.idpessoa ELSE '{}' END and p.idpessoa NOT IN ( SELECT idcli FROM mov where sitpopmov=6)".format(
            row.Cortesia, row.FirstName, row.LastName, row.Username, row.Username)
        cur.execute(sql)

    if(pd.notnull(row.DataNasc)):
        sql = "INSERT INTO DATAPESSOA (IDDTPESSOA, IDPESSOA, SITPDATA, DATA, IDINC, DHINC, IDALT, DHALT) select first 1 GEN_ID(DATAPESSOA_GEN,1), p.idpessoa, 1, '{}', 1, CURRENT_TIMESTAMP,1, CURRENT_TIMESTAMP from pessoa p left join login l on p.idpessoa=l.idlogin where p.NOMEFANTASIA like '{}' and p.NOMECOMPLETO like '{}' and l.login like CASE WHEN '{}' = '' THEN p.idpessoa ELSE '{}' END and p.idpessoa not in (select idpessoa from DATAPESSOA)".format(
            row.DataNasc, row.FirstName, row.LastName, row.Username, row.Username)
        cur.execute(sql)

    if(pd.notnull(row.RG)):
        sql = "INSERT INTO IDENTFPESS (IDIDENTFPESS, IDPESSOA, SITPIDENTF, REFERENCIA, IDINC, DHINC, IDALT, DHALT) select first 1 GEN_ID(IDENTFPESS_GEN,1), p.idpessoa, 1,'{}', 1, CURRENT_TIMESTAMP,1, CURRENT_TIMESTAMP from pessoa p left join login l on p.idpessoa=l.idlogin where p.NOMEFANTASIA like '{}' and p.NOMECOMPLETO like '{}' and l.login like CASE WHEN '{}' = '' THEN p.idpessoa ELSE '{}' END and p.idpessoa not in (select idpessoa from IDENTFPESS where sitpidentf=1)".format(
            row.RG, row.FirstName, row.LastName, row.Username, row.Username)
        cur.execute(sql)

    if(pd.notnull(row.CPF)):
        sql = "INSERT INTO IDENTFPESS (IDIDENTFPESS, IDPESSOA, SITPIDENTF, REFERENCIA, IDINC, DHINC, IDALT, DHALT) select first 1 GEN_ID(IDENTFPESS_GEN,1), p.idpessoa, 2,'{}', 1, CURRENT_TIMESTAMP,1, CURRENT_TIMESTAMP from pessoa p left join login l on p.idpessoa=l.idlogin where p.NOMEFANTASIA like '{}' and p.NOMECOMPLETO like '{}' and l.login like CASE WHEN '{}' = '' THEN p.idpessoa ELSE '{}' END and p.idpessoa not in (select idpessoa from IDENTFPESS where sitpidentf=2) and '{}' not in (select referencia from identfpess where sitpidentf=2)".format(row.CPF, row.FirstName, row.LastName, row.Username, row.Username, row.CPF)
        cur.execute(sql)

    if(pd.notnull(row['Limite Débito'])):
        sql = "update cli set LIMDEB={} where idcli=(select first 1 p.idpessoa from pessoa p left join login l on p.idpessoa=l.idlogin where p.NOMEFANTASIA like '{}' and p.NOMECOMPLETO like '{}' and l.login like CASE WHEN '{}' = '' THEN p.idpessoa ELSE '{}' END and p.idpessoa in (select idcli from cli))".format(
            row['Limite Débito'], row.FirstName, row.LastName, row.Username, row.Username)
        cur.execute(sql)

    insertPESSXFORMACNTT(cur, 1, row.Telefone, row.Cidade,
                         row.FirstName, row.LastName, row.Username)
    insertPESSXFORMACNTT(cur, 2, row.Endereço, row.Cidade,
                         row.FirstName, row.LastName, row.Username)
    insertPESSXFORMACNTT(
        cur, 3, row["E-mail"], row.Cidade, row.FirstName, row.LastName, row.Username)
    insertPESSXFORMACNTT(cur, 4, row.Celular, row.Cidade,
                         row.FirstName, row.LastName, row.Username)
    insertPESSXFORMACNTT(cur, 5, row.Responsavel, row.Cidade,
                         row.FirstName, row.LastName, row.Username)

    if(pd.notnull(row.Endereço)):
        sql = "INSERT INTO ENDERECO (IDPESSXFORMACNTT, IDBAIRRO, CEP) select first 1 IDPESSXFORMACNTT,(SELECT first 1 IDUNIDGEO from unidgeo where upper(nome)=upper('{}') and IDUNIDGEO in (select idbairro from bairro)), NULL FROM PESSXFORMACNTT WHERE IDPESSXFORMACNTT not in (select idpessxformacntt from endereco) and IDPESSOA=(select first 1 p.idpessoa from pessoa p left join login l on p.idpessoa=l.idlogin where p.NOMEFANTASIA like '{}' and p.NOMECOMPLETO like '{}' and l.login like CASE WHEN '{}' = '' THEN p.idpessoa ELSE '{}' END)".format(row.Bairro, row.FirstName, row.LastName, row.Username, row.Username)
        cur.execute(sql)

con.commit()

zip = zipfile.ZipFile('VSCyber.zip', 'w')
zip.write('VSCyber.FDB', compress_type=zipfile.ZIP_DEFLATED)

print('Importação concluída!')
