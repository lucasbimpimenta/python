# -*- coding: utf-8 -*-
import py3270
import sys
import argparse
import datetime
import pymssql 
import os
import re
import unidecode

#os.environ['TDSDUMP'] = 'stdout'

def log(string):
    f = open("log.txt", "at")
    f.write('{} - {}'.format(datetime.datetime.now(),string))
    print('{} - {}'.format(datetime.datetime.now(),string))
    f.close()

def check(emulador, linha, coluna, string):
    if emulador.string_found(linha, coluna, string):
        log('{}/{} - Esperado: {} | Encontrado: {} | OK'.format(linha,coluna,string,emulador.string_get(linha, coluna, len(string))))
        return True
    else:
        global ERROS_CHK
        ERROS_CHK.append('{}/{} - Esperado: "{}" | Encontrado: "{}"'.format(linha,coluna,string,emulador.string_get(linha, coluna, len(string))))
        log('{}/{} - Esperado: "{}" | Encontrado: "{}" | DIVERGENTE'.format(linha,coluna,string,emulador.string_get(linha, coluna, len(string))))
        return False

def estaVazio(emulador, linha, coluna, tamanho):
    if emulador.string_get(linha, coluna, tamanho).strip() == '':
        log('{}/{} - {} | VAZIO'.format(linha,coluna,emulador.string_get(linha, coluna, tamanho).strip()))
        return True
    else:
        log('{}/{} - {} | NAO VAZIO'.format(linha,coluna,emulador.string_get(linha, coluna, tamanho).strip()))
        return False

def abort(em, erro):
    global conn
    global nu_id
    atualizaReg(conn,nu_id, 'DE_LINHA_ERRO', em.string_get(23, 1, 80).strip())
    log("Abortanto o script por divergencia: {}".format(erro))
    em.terminate()
    sys.exit()

def atualizaReg(conn, nu_id, campo, valor, incremental=0):

    cursor = conn.cursor()
    if incremental == 1:
        cursor.execute("UPDATE [db7371001].[sigdu].[OS_AUTOMATICAS] SET DT_ULT_UPDATE = GETDATE(), {} = {} + CHAR(13) + CHAR(10) + %s WHERE NU_ID = %d;".format(campo, campo),(valor, nu_id))
    else:
        cursor.execute("UPDATE [db7371001].[sigdu].[OS_AUTOMATICAS] SET DT_ULT_UPDATE = GETDATE(), {} = %s WHERE NU_ID = %d;".format(campo),(valor, nu_id))
    conn.commit()
    #log('Atualizacao: Reg {} - Campo {} - Valor {}'.format(nu_id, campo, valor))

parser = argparse.ArgumentParser()
parser.add_argument("usr", help="Usuario com acesso ao SIGDU")
parser.add_argument("pwd", help="Senha do usuário com acesso ao SIGDU")
parser.add_argument("codigo", help="Código da Demanda")
parser.add_argument("simular", help="Faz somente uma simulacao", default=1, type=int, choices=[0, 1])
args = parser.parse_args()

nu_id = args.codigo
usuario = args.usr
senha = args.pwd
simulacao = args.simular

if simulacao == 1:
    print('Iniciado em modo de simulação')

#import pyodbc 
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'mg7435sr327' 
database = 'db7371001' 
username = r'gihabdv_usrsql' 
password = b'Gih4bdv\$q1@' 
#conn = pyodbc.connect('DRIVER={ODBC Driver 11 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
#print('DRIVER={SQL Server Native Client 11.0};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
#conn = pyodbc.connect('Trusted_Connection=yes; DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
##cursor = cnxn.cursor()
 
#conn = pymssql.connect(server='mg7435sr327', user=r'n6994', password=r'u6s9e9r4', database='db7371001', as_dict=True)
#conn = pymssql.connect(server=server, user=username, password=password, database=database, as_dict=True)
conn = pymssql.connect(server='mg7435sr327', user='corpcaixa\c096810', password='Au2913', database='db7371001', as_dict=True)

cursor = conn.cursor()  
cursor.execute("SELECT * FROM [db7371001].[sigdu].[OS_AUTOMATICAS] WHERE NU_ID = %d;",nu_id)
registros = cursor.fetchall()

if cursor.rowcount == 1:
    
    row = registros[0]

    if row['IC_CONFIRMACAO'] == 'P' or row['IC_CONFIRMACAO'] == 'F':

        atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Iniciando confirmacao'.format(datetime.datetime.now()))

        ERROS_CHK = list()

        em = py3270.Emulator(visible=False)

        atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Conectando ao Rede Caixa'.format(datetime.datetime.now()),1)
        em.connect('L:ibmbr10.coredf.caixa:2301')
        em.wait_for_field()

        if not check(em,17,16,'SELECIONE A OP'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Não conectei no Rede Caixa'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_CONFIRMACAO', 'F')
            abort(em, "Nao conectou no Rede Caixa")
        else:
            log("Conectei no Rede Caixa")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Conectei no Rede Caixa'.format(datetime.datetime.now()),1)

        atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Acessando o SIGDU'.format(datetime.datetime.now()),1)
        em.send_string("448")
        em.send_enter()
        em.wait_for_field()

        if not check(em,14,3,'CICS') or not check(em,24,13,'Informe Usuario') or not check(em,3,70,'RCICS'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Não acessei o SIGDU - 4.48'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_CONFIRMACAO', 'F')
            abort(em, "Nao acessou o 4.48")
        else:
            log("Aguardando usuario e senha do SIGDU")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Autenticando no SIGDU'.format(datetime.datetime.now()),1)

        em.send_string(usuario)
        em.send_string(senha)
        em.send_enter()

        em.wait_for_field()

        if not check(em,1,17,'SIGDU  GESTAO DESENVOLVIMENTO URBANO') or not check(em,2,32,'MENU PRINCIPAL') or not check(em,2,3,usuario[1:6]):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Nao autenticou no 4.48, provavelmente erro de senha'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_CONFIRMACAO', 'F')
            abort(em, "Nao autenticou no 4.48, provavelmente erro de senha")
        else:
            log("Conectei no SIGDU")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Acessei o SIGDU'.format(datetime.datetime.now()),1)

        atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Acessando o DEMA,M'.format(datetime.datetime.now()),1)
        em.move_to(22,73)
        em.send_string("EMOS,M")
        em.send_enter()
        em.wait_for_field()

        if not check(em,2,71,'EGHPO586') or not check(em,2,23,'EMOS,M') or not check(em,2,30,'CONFIRMACAO DE O.S.'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Nao acessei a tela EMOS,M, verifique se tem acesso'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_CONFIRMACAO', 'F')
            abort(em, "Nao acessei a tela DEMA,M, verifique se tem acesso")
        else:
            log("Acessei a EMOS,M")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Acessei a EMOS,M'.format(datetime.datetime.now()),1)

        #7371 7371 000300380 / 2018 01 01 01
        em.move_to(4,20)
        em.send_string(row['NU_NOVA_OS'][5:9])
        em.move_to(4,25)
        em.send_string(row['NU_NOVA_OS'][10:19])
        em.move_to(4,37)
        em.send_string(row['NU_NOVA_OS'][22:26])
        em.move_to(4,42)
        em.send_string(row['NU_NOVA_OS'][27:29])
        em.move_to(4,45)
        em.send_string(row['NU_NOVA_OS'][30:32])
        em.move_to(4,48)
        em.send_string(row['NU_NOVA_OS'][33:35])

        em.send_enter()
        em.wait_for_field()

        if not check(em,2,71,'EGHPO586') \
            or not check(em,2,23,'EMOS,M') \
            or not check(em,23,1,'INFORME OS DADOS E TECLE <ENTER>.') \
            or estaVazio(em,6,11,8) \
            or estaVazio(em,6,31,20) \
            or estaVazio(em,7,23,5) \
            or estaVazio(em,9,12,31) \
            or estaVazio(em,9,50,2) \
            or estaVazio(em,13,31,10) \
            or estaVazio(em,15,2,25):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Não consegui informar a OS Corretamente'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_CONFIRMACAO', 'F')
            abort(em, "Não consegui informar a OS Corretamente")
        else:
            log("Informei a OS Corretamente")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Informei a OS Corretamente'.format(datetime.datetime.now()),1)

        dia = datetime.datetime.now().strftime("%d")
        mes = datetime.datetime.now().strftime("%m")
        ano = datetime.datetime.now().strftime("%Y")
        dta = datetime.datetime.now().strftime("%d/%m/%Y")

        em.move_to(15,31)
        em.send_string(dta)

        em.send_enter()
        em.wait_for_field()

        if not check(em,2,71,'EGHPO586') \
            or not check(em,2,23,'EMOS,M') \
            or not check(em,23,1,'TECLE <PF2> PARA CONFIRMAR OU <PF6> PARA DESISTIR.') \
            or estaVazio(em,6,11,8) \
            or estaVazio(em,6,31,20) \
            or estaVazio(em,7,23,5) \
            or estaVazio(em,9,12,31) \
            or estaVazio(em,9,50,2) \
            or estaVazio(em,13,31,10) \
            or estaVazio(em,15,2,25) \
            or estaVazio(em,15,31,10):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Não consegui informar a data de contagem corretamente'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_CONFIRMACAO', 'F')
            abort(em, "Não consegui informar a data de contagem corretamente")
        else:
            log("Informei a data de contagem corretamente")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Informei a data de contagem corretamente'.format(datetime.datetime.now()),1)

        if simulacao == 1:
            log("Simulacao concluida - OS nao foi confirmada")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Simulacao concluida - OS nao foi confirmada'.format(datetime.datetime.now()),1)
        else:
            
            em.send_pf(2)

            if not check(em,2,71,'EGHPO586') \
                or not check(em,2,23,'EMOS,M') \
                or not check(em,23,1,'TECLE <PF6> PARA INICIAR NOVA OPERACAO.') \
                or estaVazio(em,6,11,8) \
                or estaVazio(em,6,31,20) \
                or estaVazio(em,7,23,5) \
                or estaVazio(em,9,12,31) \
                or estaVazio(em,9,50,2) \
                or estaVazio(em,13,31,10) \
                or estaVazio(em,15,2,25) \
                or estaVazio(em,15,31,10):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - Não consegui concluir a confirmacao'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_CONFIRMACAO', 'F')
                abort(em, "Não consegui concluir a confirmacao")
            else:
                log("OS confirmada com sucesso")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO_CONFIRMACAO', '{} - OS confirmada com sucesso'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_CONFIRMACAO', 'C')

        em.send_pf(12)        
        em.terminate()
    else:
        log("Demanda já executada")
else:
    log("Nenhuma demanda encontrada com o código {}".format(nu_id))

log("Fim de Execucao")

