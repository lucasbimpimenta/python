# -*- coding: utf-8 -*-
from py3270 import Emulator
import sys
import argparse
import datetime

parser = argparse.ArgumentParser()
parser.add_argument("usr", help="Usuario com acesso ao SIGDU")
parser.add_argument("pwd", help="Senha do usuário com acesso ao SIGDU")
parser.add_argument("unidade", help="Código Unidade CEF")
parser.add_argument("atv_g", help="Código da Atividade do Tipo G", type=int)
parser.add_argument("atv_inicial", help="Código da Atividade Inicial", type=int)
parser.add_argument("produto", help="Código Produto da Atividade Inicial", type=int)
parser.add_argument("linha", help="Código Linha da Atividade Inicial", type=int)
parser.add_argument("fonte", help="Código Fonte da Atividade Inicial", type=int)
parser.add_argument("objeto", help="Objeto (1 - Avaliacao/Fiscalizacao ou 2 - Analise/RAE", type=int, choices=[1, 2])

type=int, choices=[0, 1, 2]
args = parser.parse_args()

ERROS_CHK = list()

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
        ERROS_CHK.append('{}/{} - Esperado: {} | Encontrado: {}'.format(linha,coluna,string,emulador.string_get(linha, coluna, len(string))))
        log('{}/{} - Esperado: {} | Encontrado: {} | DIVERGENTE'.format(linha,coluna,string,emulador.string_get(linha, coluna, len(string))))
        return False

def abort(em, erro):
    log("Abortanto o script por divergencia: {}".format(erro))
    em.terminate()
    sys.exit()

em = Emulator(visible=False)

em.connect('L:ibmbr10.coredf.caixa:2301')
em.wait_for_field()

if not check(em,17,16,'SELECIONE A OP'):
    abort(em, "Nao conectou no Rede Caixa")
else:
    log("Conectei no Rede Caixa")

em.send_string("448")
em.send_enter()
em.wait_for_field()

if not em.string_found(14, 3, 'CICS') or not em.string_found(24, 13, 'Informe Usuario') or not em.string_found(3, 70, 'RCICS'):
    abort(em, "Nao acessou o 4.48")
else:
    log("Aguardando usuario e senha do SIGDU")

em.send_string("c096810")
em.send_string("Au2913")
em.send_enter()
em.wait_for_field()

if not em.string_found(1, 17, 'SIGDU  GESTAO DESENVOLVIMENTO URBANO'):
    abort(em, "Nao autenticou no 4.48, provavelmente erro de senha")
else:
    log("Conectei no SIGDU")

em.terminate()


