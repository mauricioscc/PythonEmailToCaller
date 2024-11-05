
from imaplib import IMAP4_SSL
from email import message_from_bytes
from time import sleep
from bs4 import BeautifulSoup
import re
import http.client
import json
import os
import sys 
import datetime
import holidays


##  PADROES

# email login data
email_cpanel_conectado = "cpanel_email@email.com"
senha_cpanel_conectado = "PASSWORD"

# delay for checking new emails
delay = 30

# Cell phone lists
business_list = "ListPath1.txt" 
non_business_list = "ListPath2.txt"

# Colors and Letters Styles

red = '\033[31m'
blue = '\033[34m'
green = '\033[32m'
yellow = '\033[33m'
purple = '\033[35m'
cyan = '\033[36m'
bold = '\033[1m'
Wbg =  '\033[40m'
end = '\033[m'

print(f""" {blue}   Email to Caller Python System{end}
 \n{yellow} --> {delay} delay em segundos{end}   
 {green}--> Logged in {email_cpanel_conectado}{end}

 {cyan}--> Lista de horários utéis
 --> Lista de horários não úteis {end}

 {Wbg}{purple}{bold}--> Version 2.0.0 (05/11/2024){end}

 {blue}---------------------------------------------------------{end}""")

while True:
    try:

        IMAP_SERVER = 'endereço do servidor IMAP'
        
        USERNAME = email_cpanel_conectado
        PASSWORD =  senha_cpanel_conectado

        remetente_esperado = 'email do remetente'

        mail = IMAP4_SSL(IMAP_SERVER, 993)
        mail.login(USERNAME, PASSWORD)

        mail.select('inbox')

        result, data = mail.search(None, "UNSEEN")

        for num in data[0].split():
            result, data = mail.fetch(num, "(RFC822)")
            raw_email = data[0][1]

            msg = message_from_bytes(raw_email)

            remetente = msg["From"]
            corpo = ""

            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    if "attachment" not in content_disposition:
                        payload = part.get_payload(decode=True)
                        if payload is not None:
                            corpo_parte = payload.decode('utf-8', 'ignore')
                            corpo += corpo_parte if corpo_parte is not None else ""
            else:
                corpo = msg.get_payload(decode=True).decode('utf-8', 'ignore')

            if remetente != remetente_esperado:
                print(f'\nRemetente diferente do previsto\n--> Remetente esperado: {remetente_esperado}\n--> Remetente recebido: {remetente}\n')
            
            # Remover prefixo variável do corpo do e-mail usando "re" 
            padrao_prefixo = re.compile(r'Olá.*?porObservações', re.DOTALL) 
            match = re.search(padrao_prefixo, corpo)
            if match:
                corpo = corpo.replace(match.group(0), '').strip()
            padrao_remover = re.compile(r'\*www\.sitrad\.com\.br/ <http://www\.sitrad\.com\.br/>.*?<html>', re.DOTALL)
            padrao_remover_html = re.compile(r'<body><h1>.*?</html>', re.DOTALL)
            corpo = re.sub(padrao_remover, '', corpo)
            corpo = re.sub(padrao_remover_html, "", corpo)

            # Verificar se o corpo parece ser uma marcacao HTML válida
            if corpo.startswith("<") and corpo.endswith(">"):
                soup = BeautifulSoup(corpo, 'html.parser')
                texto_sem_html = soup.get_text()
                linhas_corpo = texto_sem_html.split('\n')
            else:
                linhas_corpo = corpo.split("\n")
            imprimiu = False

            # BUSCA E FILTRO DE PALAVRAS ESPECIFICAS
            Rec = 'Recorrente'
            Nov = 'Novo'
            Fin = 'Finalizado'

            if Fin in corpo:
                index_finalizado = corpo.index(Fin)
                corpo = corpo[:index_finalizado].strip()
            
            for linha in linhas_corpo: 
                linha_strip = linha.strip()
                
                if linha_strip.startswith(Nov) and not imprimiu:
                    tipo_alarme = "Novo"
                    print(yellow, "--" * 40, end)
                    print(green, "Email to Caller Python System", end)
                    print("\n")
                    print(corpo)
                    print(yellow, "--" * 40, end)
                    imprimiu = True

                elif linha_strip.startswith(Rec) and not imprimiu:
                    tipo_alarme = "Recorrente"
                    print(yellow, "--" * 40, end)
                    print(green, "Email to Caller Python System", end)
                    print("\n")
                    print(corpo)
                    print(yellow, "--" * 40, end)
                    imprimiu = True
 
                    continue

           ## RE-FORMATACAO DO CORPO!!!
            if "Valor:" in corpo:
                corpo_TTS = re.sub(r'(Valor:)', r'\\pause=500 \1 \\pause=250', corpo)
            if Nov in corpo_TTS:
                corpo_TTS = re.sub(r'(Novo)', r'\\pause=500 \\speed=50 \1 \\pause=250', corpo_TTS)
            if Rec in corpo_TTS:
                corpo_TTS = re.sub(r'(Recorrente)', r'\\pause=500 \\speed=50 \1 \\pause=250', corpo_TTS)
            if "Limites" in corpo:
                corpo_TTS = re.sub(r'(Limites:)', r'\\pause=500 \1 \\pause=250', corpo_TTS)

            re_mensagem = (f" \pause=1800 Email to Caller Python System {corpo_TTS}")
            mensagem = re_mensagem

            diretorio_script = os.path.dirname(os.path.abspath(sys.argv[0]))            
            diretorio_atual = os.path.dirname(os.path.abspath(__file__))
            caminho_lista = os.path.join(diretorio_script, business_list) # LISTA UTIL
            caminho_lista_sec = os.path.join(diretorio_script, non_business_list) # LISTA NAO UTIL
            
            if os.path.exists(caminho_lista):
                with open(caminho_lista, 'r') as file:
                    lista = [line.strip() for line in file]
            else:
                caminho_lista_temporario = os.path.join(sys._MEIPASS, business_list)
                with open(caminho_lista_temporario, 'r') as file:
                    lista = [line.strip() for line in file]


            if os.path.exists(caminho_lista_sec):
                with open(caminho_lista_sec, 'r') as file:
                    lista_sec = [line.strip() for line in file]
            else:
                caminho_lista_temporario_sec = os.path.join(sys._MEIPASS, non_business_list)
                with open(caminho_lista_temporario_sec, 'r') as file:
                    lista_sec = [line.strip() for line in file]

            dt_now = datetime.datetime.now()
            data_atual = dt_now.date()
            dia_semana = dt_now.weekday() # 0 - segunda feira & 6 - domingo
            mes_atual = dt_now.month
            hora = dt_now.hour
            feriados_Brasil_RJ = holidays.country_holidays('BR', subdiv='RJ')

            def dia_util(data_atual, feriados_Brasil_RJ):
                if data_atual not in feriados_Brasil_RJ:
                    return True # SIM, HOJE É UM DIA ÚTIL
                else:
                    return False # NÃO, HOJE NÃO É UM DIA ÚTIL     
                      
            def util(dia_semana, hora):
                if dia_semana < 5 and hora >= 7 and hora < 17:
                   return True
                else:
                    return False

            if dia_util(data_atual, feriados_Brasil_RJ) and util(dia_semana, hora): # dentro do tempo útil
                print(f"Lista de telefones dentro do horário útil em uso ({business_list}):")
                for contact in lista:
                                print(f"ligando para {contact}")
                                payload1 = json.dumps({
                                                "dest": contact, 
                                                "text": mensagem,
                                                "voice": "MS|pt-BR-FranciscaNeural",
                                                "name": "Email to Caller Python System",
                                                "enconding": "UTF-8"
                                                })
                                conn = http.client.HTTPSConnection("vox.velip.com.br")
                                headers = {
                                    'Content-Type': 'application/json',
                                    'Authorization': 'VELIP Authorization code'
                                    }
                                conn.request("POST", "/api/v2/MakeTTSCall", payload1, headers)
                                res = conn.getresponse()
                                data = res.read()
                                print(data.decode("utf-8"))

            elif not dia_util(data_atual, feriados_Brasil_RJ) or not util(dia_semana, hora): # fora do tempo útil
                print(f"Lista de telefones fora do horário útil em uso ({non_business_list}):")
                for contact in lista_sec:
                                print(f"ligando para {contact}")
                                payload1 = json.dumps({
                                                "dest": contact, 
                                                "text": mensagem,
                                                "voice": "MS|pt-BR-FranciscaNeural",
                                                "name": "Email to Caller Python System",
                                                "enconding": "UTF-8"
                                                })
                                conn = http.client.HTTPSConnection("vox.velip.com.br")
                                headers = {
                                    'Content-Type': 'application/json',
                                    'Authorization': 'VELIP Authorization code'
                                    }
                                conn.request("POST", "/api/v2/MakeTTSCall", payload1, headers)
                                res = conn.getresponse()
                                data = res.read()
                                print(data.decode("utf-8"))

            else: # erro 
                print(f'{bold}{red}ERRO NA CONFIGURACAO DE FERIADOS, FINS DE SEMANA E DIAS UTEIS \n Nenhuma ligacao foi EFETUADA !!!{end}')                    

        # FECHANDO A CONEXÃO COM O SERVIDOR IMAP
        mail.close()
        mail.logout()

        # Aguardar X segundos antes da próxima verificação
        sleep(delay)

    except Exception as e:
        print("Ocorreu um erro:", e)

