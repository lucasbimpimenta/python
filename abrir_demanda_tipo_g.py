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

    if row['IC_EXECUCAO'] == 'P':

        atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Iniciando processamento'.format(datetime.datetime.now()))

        ERROS_CHK = list()

        em = py3270.Emulator(visible=False)

        atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Conectando ao Rede Caixa'.format(datetime.datetime.now()),1)
        em.connect('L:ibmbr10.coredf.caixa:2301')
        em.wait_for_field()

        if not check(em,17,16,'SELECIONE A OP'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Não conectei no Rede Caixa'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Nao conectou no Rede Caixa")
        else:
            log("Conectei no Rede Caixa")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Conectei no Rede Caixa'.format(datetime.datetime.now()),1)

        atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Acessando o SIGDU'.format(datetime.datetime.now()),1)
        em.send_string("448")
        em.send_enter()
        em.wait_for_field()

        if not check(em,14,3,'CICS') or not check(em,24,13,'Informe Usuario') or not check(em,3,70,'RCICS'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Não acessei o SIGDU - 4.48'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Nao acessou o 4.48")
        else:
            log("Aguardando usuario e senha do SIGDU")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Autenticando no SIGDU'.format(datetime.datetime.now()),1)

        em.send_string(usuario)
        em.send_string(senha)
        em.send_enter()

        em.wait_for_field()

        if not check(em,1,17,'SIGDU  GESTAO DESENVOLVIMENTO URBANO') or not check(em,2,32,'MENU PRINCIPAL') or not check(em,2,3,usuario[1:6]):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Nao autenticou no 4.48, provavelmente erro de senha'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Nao autenticou no 4.48, provavelmente erro de senha")
        else:
            log("Conectei no SIGDU")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Acessei o SIGDU'.format(datetime.datetime.now()),1)

        atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Acessando o DEMA,M'.format(datetime.datetime.now()),1)
        em.move_to(22,73)
        em.send_string("DEMA,M")
        em.send_enter()
        em.wait_for_field()

        if not check(em,2,71,'EGHPO571') or not check(em,2,23,'DEMA,M') or not check(em,23,1,'INFORME A OPCAO E TECLE <ENTER>'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Nao acessei a tela DEMA,M, verifique se tem acesso'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Nao acessei a tela DEMA,M, verifique se tem acesso")
        else:
            log("Acessei a DEMA,M")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Acessei a DEMA,M'.format(datetime.datetime.now()),1)

        em.send_pf(10)
        em.wait_for_field()

        if not check(em,16,6,'Codigo Unidade CEF') or not check(em,2,71,'EGHPO571') or not check(em,2,23,'DEMA,M'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Nao acessei o F10 da DEMA,M, verifique se tem acesso'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Nao acessei o F10 da DEMA,M, verifique se tem acesso")
        else:
            log("Acessei ao F10 da DEMA,M")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Acessei o F10 da DEMA,M'.format(datetime.datetime.now()),1)

        em.move_to(16,26)
        em.send_string(row['CP_UNIDADE'])
        em.send_enter()
        em.wait_for_field()

        if not check(em,16,26,row['CP_UNIDADE']) or not check(em,2,71,'EGHPO571') or not check(em,23,1,'TECLE <ENTER> PARA PROSSEGUIR COM A DEMANDA') or estaVazio(em,16,33,45):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar unidade no F10 da DEMA,M'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao informar unidade no F10 da DEMA,M")
        else:
            log("Unidade informada no F10 da DEMA,M")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Unidade informada no F10 da DEMA,M'.format(datetime.datetime.now()),1)

        em.send_enter()
        em.wait_for_field()

        if not check(em,4,2,'Deseja fazer demanda para atividade do grupo G') or not check(em,2,71,'EGHPO571') or not check(em,23,1,'INFORME OS DADOS NECESSARIOS E PRESSIONE <ENTER>'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar unidade no F10 da DEMA,M (segundo <Enter>)'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao informar unidade no F10 da DEMA,M (segundo <Enter>)")
        else:
            log("Unidade informada no F10 da DEMA,M (segundo <Enter>)")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Unidade informada no F10 da DEMA,M (segundo <Enter>)'.format(datetime.datetime.now()),1)

        em.move_to(4,51)
        em.send_string('S')
        em.send_enter()
        em.wait_for_field()

        if not check(em,4,51,'S') or not check(em,4,2,'Deseja fazer demanda para atividade do grupo G') or not check(em,2,71,'EGHPO571') or not check(em,5,2,'Atividade para o grupo G'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao selecionar demanda Tipo G'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao selecionar demanda Tipo G")
        else:
            log("Selecionou demanda Tipo G")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Selecionou demanda Tipo G'.format(datetime.datetime.now()),1)

        em.move_to(5,30)
        em.send_string('G')
        em.move_to(5,32)
        em.send_string(str(row['CP_NUM_TIPO_G']).zfill(3))
        em.send_enter()
        em.wait_for_field()

        if not estaVazio(em,23,1,45) or not check(em,4,2,'Deseja fazer demanda para atividade do grupo G') or not check(em,2,71,'EGHPO571') or not check(em,5,2,'Atividade para o grupo G'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar a nova demanda Tipo G'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao informar a nova demanda Tipo G")
        else:
            log("Informou nova demanda Tipo G")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Informou nova demanda Tipo G'.format(datetime.datetime.now()),1)

        em.move_to(8,15)
        em.send_string(row['CP_ATV_LETRA'])
        em.move_to(8,17)
        em.send_string(str(row['CP_ATV_NUM']).zfill(3))
        em.move_to(9,16)
        em.send_string(str(row['CP_PRODUTO']).zfill(3))
        em.move_to(10,16)
        em.send_string(str(row['CP_LINHA']).zfill(3))
        em.move_to(11,16)
        em.send_string(str(row['CP_FONTE']).zfill(3))

        em.send_enter()
        em.wait_for_field()

        if not estaVazio(em,23,1,45) or estaVazio(em,8,23,45) or estaVazio(em,9,23,45) or estaVazio(em,10,23,45) or estaVazio(em,11,23,45) or not check(em,4,2,'Deseja fazer demanda para atividade do grupo G') or not check(em,2,71,'EGHPO571') or not check(em,13,2,'O servico sera executado por terceiro ou empregado Caixa'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar demanda antiga, produto, linha ou fonte'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao informar demanda antiga, produto, linha ou fonte")
        else:
            log("Informou demanda antiga, produto, linha e fonte")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Informou demanda antiga, produto, linha e fonte'.format(datetime.datetime.now()),1)

        em.move_to(13,61)
        em.send_string('T')
        em.send_enter()
        em.wait_for_field()

        if not estaVazio(em,23,1,45) or not check(em,13,61,'T') or not check(em,4,2,'Deseja fazer demanda para atividade do grupo G') or not check(em,2,71,'EGHPO571') or not check(em,5,2,'Atividade para o grupo G'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao Informou demanda sera terceirizada'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao informar Informou demanda sera terceirizada")
        else:
            log("Informou demanda sera terceirizada")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Informou demanda sera terceirizada'.format(datetime.datetime.now()),1)

        em.move_to(15,15)
        em.send_string(str(row['CP_OBJETO']))
        em.send_enter()
        em.wait_for_field()

        if row['CP_OBJETO'] == 1:

            if not check(em,2,72,'EGHPO573') or not check(em,2,30,'IMOVEL') or not check(em,23,1,'INFORME O SISTEMA E TECLE <ENTER>') or not check(em,4,17,'SIAPF') or not check(em,7,3,'Comarca'):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar objeto'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar objeto")
            else:
                log("Informou objeto do Tipo 1")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Informou objeto do Tipo 1'.format(datetime.datetime.now()),1)

            em.move_to(4,12)
            em.send_string('3')
            em.send_enter()
            em.wait_for_field()

            if not check(em,2,72,'EGHPO573') or not check(em,2,30,'IMOVEL') or not check(em,4,63,'S/N______________') or not check(em,23,1,'INFORME A IDENTIFICACAO NO CARTORIO DO IMOVEL E TECLE <ENTER>'):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar o sistema Outros'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar o sistema Outros")
            else:
                log("Informou sistema Outros")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Informou sistema Outros'.format(datetime.datetime.now()),1)

            em.move_to(6,22)
            em.send_string(row['CP_MATRICULA'])

            em.move_to(6,70)
            em.send_string(row['CP_OFICIO'])

            em.move_to(7,22)
            em.send_string(row['CP_COMARCA'])

            em.move_to(7,70)
            em.send_string(row['CP_UF_COMARCA'])

            em.send_enter()
            em.wait_for_field()

            if not check(em,20,54,'TECLE <ENT> P/ CONTINUAR') \
            or estaVazio(em,11,9,5) \
            or estaVazio(em,11,27,2) \
            or estaVazio(em,11,45,5) \
            or estaVazio(em,14,17,15) \
            or estaVazio(em,14,40,40) \
            or estaVazio(em,15,17,20) \
            or estaVazio(em,19,40,40):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar a matricula'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar a matricula")
            else:
                log("Confirmou o Matricula")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Confirmou a Matricula'.format(datetime.datetime.now()),1)

            em.send_enter()
            em.wait_for_field()

            if not estaVazio(em,20,54,24) or not check(em,23,1,"INFORME O CEP DO ENDERECO E TECLE <ENTER>"):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao confirmar a Matricula (segundo enter)'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao confirmar a Matricula (segundo enter)")
            else:
                log("Confirmou o Matricula (segundo enter)")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Confirmou o Matricula (segundo enter)'.format(datetime.datetime.now()),1)

            em.move_to(11,9)
            em.send_string(row['CP_CEP'][0:5])
            em.move_to(11,17)
            em.send_string(row['CP_CEP'][5:8])
            em.send_enter()
            em.wait_for_field()

            if not estaVazio(em,23,1,45) \
            or not check(em,2,72,'EGHPO573') \
            or not check(em,2,30,'IMOVEL') \
            or not check(em,4,63,'S/N______________') \
            or estaVazio(em,11,27,2) \
            or estaVazio(em,11,45,5) \
            or estaVazio(em,11,69,2) \
            or estaVazio(em,12,33,20) \
            or estaVazio(em,12,60,18):                
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar o CEP'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar o CEP")
            else:
                log("CEP informado")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - CEP informado'.format(datetime.datetime.now()),1)

            em.send_enter()
            em.wait_for_field()

            if not check(em,23,1,'INFORME O ENDERECO E TECLE <ENTER>') \
            or not check(em,2,72,'EGHPO573') \
            or not check(em,2,30,'IMOVEL') \
            or not check(em,4,63,'S/N______________') \
            or estaVazio(em,11,27,2) \
            or estaVazio(em,11,45,5) \
            or estaVazio(em,11,69,2) \
            or estaVazio(em,12,33,20) \
            or estaVazio(em,12,60,18):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar o CEP (segundo <enter>)'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar o CEP (segundo <enter>)")
            else:
                log("CEP informado (segundo <enter>)")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - CEP informado (segundo <enter>)'.format(datetime.datetime.now()),1)

            em.move_to(14,17)
            em.send_string(unidecode.unidecode(row['CP_LOG_TIPO1']))

            em.move_to(14,40)
            em.send_string(unidecode.unidecode(row['CP_LOG_NOME']))

            em.move_to(15,17)
            em.send_string(row['CP_LOG_NUM1'])

            if(row['CP_LOG_NUM2'] is None or row['CP_LOG_NUM2'].strip() == ''):
                row['CP_LOG_NUM2'] = '      '

            em.move_to(19,17)
            em.send_string(row['CP_LOG_NUM2'])

            em.move_to(19,40)
            em.send_string(unidecode.unidecode(row['CP_LOG_BAIRRO']))

            em.send_enter()
            em.wait_for_field()

########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################

        elif row['CP_OBJETO'] == 2:

            if not estaVazio(em,23,1,45) or not check(em,2,71,'EGHPO571') or not check(em,15,12,2):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar objeto'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar objeto")
            else:
                log("Confirmou o Objeto")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Confirmou o Objeto'.format(datetime.datetime.now()),1)

            em.send_enter()
            em.wait_for_field()

            if not check(em,2,72,'EGHPO574') or not check(em,2,30,'EMPREENDIMENTO') or not check(em,23,1,'INFORME OS DADOS NECESSARIOS E PRESSIONE <ENTER>') or not check(em,6,25,'SIAPF'):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar objeto (segundo <enter>)'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar objeto (segundo <enter>)")
            else:
                log("Informou objeto do Tipo 2 (segundo <enter>)")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Informou objeto do Tipo 2 (segundo <enter>)'.format(datetime.datetime.now()),1)

            em.move_to(6,20)
            em.send_string('3')
            em.send_enter()
            em.wait_for_field()

            if not check(em,2,72,'EGHPO574') or not check(em,2,30,'EMPREENDIMENTO') or not check(em,6,57,'S/N______________') or not check(em,6,25,'SIAPF'):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar o sistema Outros'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar o sistema Outros")
            else:
                log("Informou sistema Outros")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Informou sistema Outros'.format(datetime.datetime.now()),1)

            em.move_to(12,9)
            em.send_string(row['CP_CEP'])
            em.send_enter()
            em.wait_for_field()

            if not estaVazio(em,23,1,45) or not check(em,2,72,'EGHPO574') or not check(em,2,30,'EMPREENDIMENTO') or not check(em,6,57,'S/N______________') or estaVazio(em,12,27,2) or estaVazio(em,12,45,5) or estaVazio(em,12,69,2) or estaVazio(em,13,33,20) or estaVazio(em,13,60,18):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar o CEP'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar o CEP")
            else:
                log("CEP informado")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - CEP informado'.format(datetime.datetime.now()),1)

            em.send_enter()
            em.wait_for_field()

            if not check(em,23,1,'INFORME OS DADOS NECESSARIOS E PRESSIONE') or not check(em,2,72,'EGHPO574') or not check(em,2,30,'EMPREENDIMENTO') or not check(em,6,57,'S/N______________') or estaVazio(em,12,27,2) or estaVazio(em,12,45,5) or estaVazio(em,12,69,2) or estaVazio(em,13,33,20) or estaVazio(em,13,60,18):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar o CEP (segundo <enter>)'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao informar o CEP (segundo <enter>)")
            else:
                log("CEP informado (segundo <enter>)")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - CEP informado (segundo <enter>)'.format(datetime.datetime.now()),1)

            em.move_to(15,17)
            em.send_string(unidecode.unidecode(row['CP_LOG_TIPO1']))

            em.move_to(15,40)
            em.send_string(unidecode.unidecode(row['CP_LOG_NOME']))

            em.move_to(16,17)
            em.send_string(row['CP_LOG_NUM1'])

            em.move_to(17,17)
            em.send_string(row['CP_LOG_TIPO2'])

            em.move_to(18,17)
            em.send_string(row['CP_LOG_NUM2'])

            em.move_to(18,40)
            em.send_string(unidecode.unidecode(row['CP_LOG_BAIRRO']))

            em.send_enter()
            em.wait_for_field()

        else:
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar objeto. Valor não previsto (1 ou 2)'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao informar objeto. Valor não previsto (1 ou 2)")

        ### TELA DE CONVOCACAO

        if not check(em,2,71,'EGHPO576') \
        or not check(em,2,30,'CONVOCACAO') \
        or not check(em,4,2,'Demanda/OS criada:') \
        or not estaVazio(em,4,21,20) \
        or not check(em,7,13,row['CP_ATV_LETRA'].strip()) \
        or not check(em,7,15,str(row['CP_ATV_NUM']).strip().zfill(3)) \
        or not check(em,7,19,'G') \
        or not check(em,7,20,str(row['CP_NUM_TIPO_G']).strip().zfill(3)):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao informar dados o Imóvel'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao informar dados o Imóvel")
        else:
            log("Dados do Imóvel informado")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Dados do Imóvel informado'.format(datetime.datetime.now()),1)

        em.send_pf(9)
        em.wait_for_field()

        if not check(em,2,71,'EGHPO577') \
        or not check(em,2,30,'ESCOLHA DIRIGIDA EMPRESA') \
        or not check(em,6,5,'Empresa'):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao Acessar a Escolha Dirigida'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao Acessar a Escolha Dirigida")
        else:
            log("Acessou Escolha Dirigida")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Acessou Escolha Dirigida'.format(datetime.datetime.now()),1)

        em.move_to(6,15)
        em.send_string(row['CP_CNPJ'][0:8])

        em.move_to(6,26)
        em.send_string(row['CP_CNPJ'][8:12])

        em.move_to(6,33)
        em.send_string(row['CP_CNPJ'][12:14])

        em.move_to(18,19)
        em.send_string(unidecode.unidecode(row['CP_JUSTIFICATIVA']))

        em.send_enter()
        em.wait_for_field()

        if not check(em,2,71,'EGHPO576') \
        or not check(em,2,30,'CONVOCACAO') \
        or not check(em,4,2,'Demanda/OS criada:') \
        or estaVazio(em,5,28,16) \
        or not check(em,5,9,row['CP_CNPJ'][0:8]) \
        or not check(em,5,18,row['CP_CNPJ'][8:12]) \
        or not check(em,5,23,row['CP_CNPJ'][12:14]):
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao fazer Escolha Dirigida da Empresa'.format(datetime.datetime.now()),1)
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
            abort(em, "Erro ao fazer Escolha Dirigida da Empresa")
        else:
            log("Escolha Dirigida da Empresa com sucesso")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Escolha Dirigida da Empresa com sucesso'.format(datetime.datetime.now()),1)

        em.move_to(7,62)
        em.send_string(row['CP_OS_ORIG_UND_DEB'])

        em.move_to(8,68)
        em.send_string(row['CP_VR_PAGAR'])

        if(row['CP_VR_DESLOCAMENTO'] is None or row['CP_VR_DESLOCAMENTO'].strip() == ''):
            row['CP_VR_DESLOCAMENTO'] = '            '

        em.move_to(9,28)
        em.send_string(row['CP_VR_DESLOCAMENTO'])

        em.move_to(10,28)
        em.send_string('1')

        em.move_to(11,16)
        em.send_string(unidecode.unidecode(row['CP_CONTATO']))

        em.move_to(12,20)
        em.send_string(unidecode.unidecode(row['CP_TELEFONE']))

        em.move_to(13,31)
        em.send_string(unidecode.unidecode(row['CP_LOCAL_RETIRADA']))

        em.move_to(16,14)
        em.send_string(unidecode.unidecode(row['CP_OBSERVACOES']))

        if simulacao == 1:
            log("Simulacao concluida - OS nao foi aberta")
            atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Simulacao concluida - OS nao foi aberta'.format(datetime.datetime.now()),1)
        else:

            em.send_enter()
            em.wait_for_field()

            if not check(em,2,71,'EGHPO576') \
            or not check(em,2,30,'CONVOCACAO') \
            or not check(em,4,2,'Demanda/OS criada:') \
            or estaVazio(em,5,28,16) \
            or not check(em,23,1,'TECLE <PF2> PARA CONFIRMAR OU <PF6> PARA DESISTIR.'):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Dados para abertura preenchidos com erro'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao Abrir OS")
            else:
                log("Dados para abertura preenchidos com sucesso")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Dados para abertura preenchidos com sucesso'.format(datetime.datetime.now()),1)

            em.send_pf(2)
            em.wait_for_field()

            print(em.string_get(23, 1, 35).strip())
            print(em.string_get(4, 21, 19).strip())
            print(em.string_get(4, 43, 11).strip())

            atualizaReg(conn,nu_id, 'NU_NOVA_OS', em.string_get(4, 21, 35).strip())

            if not check(em,2,71,'EGHPO576') \
            or not check(em,2,30,'CONVOCACAO') \
            or not check(em,4,2,'Demanda/OS criada:') \
            or estaVazio(em,4,21,19) \
            or estaVazio(em,4,43,11) \
            or not check(em,23,1,'E-mail enviado com sucesso.'):
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - Erro ao Abrir OS'.format(datetime.datetime.now()),1)
                atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'F')
                abort(em, "Erro ao Abrir OS")
            else:
                log("OS Aberta com sucesso")
                atualizaReg(conn,nu_id, 'DE_PROCESSAMENTO', '{} - OS Aberta com sucesso'.format(datetime.datetime.now()),1)

            atualizaReg(conn,nu_id, 'NU_NOVA_OS', em.string_get(4, 21, 35).strip())
            atualizaReg(conn,nu_id, 'IC_EXECUCAO', 'C')

            #os.system("python confirmar_os.py {} {} {} {}".format(usuario, senha, nu_id, simulacao))

        em.send_pf(12)
        em.terminate()
    else:
        log("Demanda já executada")
else:
    log("Nenhuma demanda encontrada com o código {}".format(nu_id))

log("Fim de Execucao")

