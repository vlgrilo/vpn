from dateutil.relativedelta import relativedelta
from datetime import datetime
import requests
import gspread
import time

#CREDENCIAIS GOOGLE CLOUD 
gc = gspread.service_account(filename='service_account.json')
time.sleep(1)

#CHAVE PLANILHA
sh = gc.open_by_key('CHAVE PANILHA')
time.sleep(1)

#PAGINA POR TITULO 
ws = sh.worksheet("VPN'S")
time.sleep(1)

#TOKEN ULTILIZADO PARA MANIPULAR O BOT
token = 'TOKEN'

#ID DO CHAT ONDE SERA ENVIADO AS MENSAGENS
chat_id = 'ID CHAT TELEGRAN'

#ENVIA MENSAGEM ULTILIZANDO BOT EM GRUPO ESPECIFICO
def send_message(token, chat_id, msg):
    try:
        data = {"chat_id": chat_id, "text": msg}
        url = "https://api.telegram.org/bot{}/sendMessage".format(token)
        requests.post(url, data)
    except Exception as e:
        print("Erro no sendMessage:", e)

#CONTA LINHAS PLANILHA
linhas = ws.row_count
time.sleep(1)

#VARIAVEIS DE APOIO
contlinha = 0
contador2 = 0
nada = 0

#PERCORRE A PLANILHA
while contlinha < linhas:
    time.sleep(1)
    contlinha += 1

    #PEGA LINHAS COM "SESRVIDOR PUBLICO" OU "SERVIDOR PÚBLICO"
    if(ws.cell(contlinha, 7).value == 'Servidor Publico' or ws.cell(contlinha, 7).value == 'Servidor Público'):
        time.sleep(1)
        contador2 = 0
        str_vencimento = ws.cell(contlinha, 1).value
        time.sleep(1)
        vencimento = datetime.strptime(str_vencimento, '%d/%m/%Y').date()
        hoje = datetime.today()

        #ANO VPN MENOR QUE ANO ATUAL
        if (vencimento.year < hoje.year):
            addano = hoje.year - vencimento.year
            vencimento = vencimento + relativedelta(years = addano)

        #DIFERENÇA DE DIAS
        diff = relativedelta(vencimento , hoje)

        #VENCIMENTO IGUAL DIA ATUAL
        if (diff.months == 0 and diff.days == 0 and diff.hours < 0):
            nada = 1
            msg = ("-----------------------------------\nVPN VENCENDO EM 0 DIAS \nNome: " + str(ws.cell(contlinha, 2).value) + "\nLogin: " + str(ws.cell(contlinha, 3).value) + "\nUID (NIS): " + str(ws.cell(contlinha, 4).value) + "\nUID(TUPA): " + str(ws.cell(contlinha, 5).value) + "\nEmail: " + str(ws.cell(contlinha, 6).value) + "\nVinculo: " + str(ws.cell(contlinha, 7).value) + "\nServidor: " + str(ws.cell(contlinha, 8).value) + "\nMac: " + str(ws.cell(contlinha, 9).value) + "\n-----------------------------------")
            send_message(token, chat_id, msg)
            time.sleep(2)

        #VENCIMENTO IGUAL AMANHA
        if (diff.months == 0 and diff.days == 0 and diff.hours >= 0):
            nada = 1
            msg = ("-----------------------------------\nVPN VENCENDO EM 1 DIA \nNome: " + str(ws.cell(contlinha, 2).value) + "\nLogin: " + str(ws.cell(contlinha, 3).value) + "\nUID (NIS): " + str(ws.cell(contlinha, 4).value) + "\nUID(TUPA): " + str(ws.cell(contlinha, 5).value) + "\nEmail: " + str(ws.cell(contlinha, 6).value) + "\nVinculo: " + str(ws.cell(contlinha, 7).value) + "\nServidor: " + str(ws.cell(contlinha, 8).value) + "\nMac: " + str(ws.cell(contlinha, 9).value) + "\n-----------------------------------")
            send_message(token, chat_id, msg)
            time.sleep(2)

        #VENCIMENTO ENTRE 1 E 7 DIAS
        if (diff.months == 0 and diff.days > 0 and diff.days < 7):
            nada = 1
            msg = ("-----------------------------------\nVPN VENCENDO EM " + str(diff.days + 1) + " DIAS \nNome: " + str(ws.cell(contlinha, 2).value) + "\nLogin: " + str(ws.cell(contlinha, 3).value) + "\nUID (NIS): " + str(ws.cell(contlinha, 4).value) + "\nUID(TUPA): " + str(ws.cell(contlinha, 5).value) + "\nEmail: " + str(ws.cell(contlinha, 6).value) + "\nVinculo: " + str(ws.cell(contlinha, 7).value) + "\nServidor: " + str(ws.cell(contlinha, 8).value) + "\nMac: " + str(ws.cell(contlinha, 9).value) + "\n-----------------------------------")
            send_message(token, chat_id, msg)
            time.sleep(2)

    #PEGA LINHAS VAZIAS
    elif (ws.cell(contlinha, 7).value) == None:
        contador2 += 1
        #PARA CODIGO CASO TENHA  DE 5 LINHAS VAZIAS
        if(contador2 == 5):
            if(nada == 0):
                msg = ("-----------------------------------\nNENHUMA VPN VENCENDO NA SEMANA\n-----------------------------------")
                send_message(token, chat_id, msg)
                exit()
            else:
                exit()
        time.sleep(2)

    #OUTRAS LINHAS (BOLSISTA, TERCEIRO , ETC)
    else:
        contador2 = 0
        time.sleep(2)