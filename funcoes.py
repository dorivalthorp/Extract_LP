# Imports
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
from email.mime.text import MIMEText
import smtplib

from zipfile import ZipFile, ZIP_DEFLATED
from winotify import Notification, audio
from datetime import datetime, timedelta
from sqlalchemy import create_engine
import win32com.client as w32
from bs4 import BeautifulSoup
import logging as log
from tqdm import tqdm
import pandas as pd
import requests
import psycopg2
import zipfile
import imaplib
import shutil
import gnupg
import email
import json
import os

# Obter a data e hora atual no formato AAAA_MM_DD_hhmm para o processo inteiro
data_hora_atual = datetime.now().strftime('%Y_%m_%d_%H%M')

# LOG
log.basicConfig(filename=f'liveperson_{data_hora_atual}.log', level=log.WARNING, format='%(asctime)s - %(message)s')
http_logger = log.getLogger('werkzeug')
http_logger.setLevel(log.ERROR)

log.warning('Iniciando API')

def notificacao(nome_app,titulo,icone,execucao):
    notificacao = Notification(app_id=nome_app, title=titulo,icon=icone)
    notificacao.set_audio(audio.LoopingAlarm,loop=False)
    notificacao.add_actions(label='Clique Aqui',launch=execucao)
    notificacao.show()

def ler_config():
    try:
        # Lendo as configurações do arquivo config.json
        with open("config.json") as config_file:
            config = json.load(config_file)

        email_user = config["email_user"]
        email_password = config["email_password"]
        mail_server = config["mail_server"]
        mail_port = config["mail_port"]

        return email_user, email_password, mail_server, mail_port
    except FileNotFoundError:
        print("Arquivo config.json não encontrado.")
        return 'ERRO', 'ERRO', 'ERRO', 'ERRO'
    except json.JSONDecodeError:
        print("Erro ao decodificar o arquivo config.json.")
        return 'ERRO', 'ERRO', 'ERRO', 'ERRO'
    except KeyError as e:
        print("Chave ausente no arquivo config.json:", e)
        return 'ERRO', 'ERRO', 'ERRO', 'ERRO'
    except Exception as e:
        print("Erro inesperado:", e)
        return 'ERRO', 'ERRO', 'ERRO', 'ERRO'

def pega_link_email(email_user, email_password, mail_server, mail_port):
    try:
        temporary_url='ERRO'

        # Conectando ao servidor IMAP
        mail = imaplib.IMAP4_SSL(mail_server, mail_port)
        mail.login(email_user, email_password)
        mail.select("Getnet")  # Seleciona a caixa de entrada

        # Obtendo a data atual
        current_date = datetime.now().strftime("%d-%b-%Y")  # Formato: dia-mês-ano

        # Definindo critério de pesquisa para a data atual e assunto específico
        search_criteria = '(ON {0} SUBJECT "Data Transporter")'.format(current_date)

        # Pesquisa por emails com base no critério
        status, email_ids = mail.search(None, search_criteria)

        # Recuperando os IDs dos emails
        email_id_list = email_ids[0].split()

        # Loop para processar cada email
        for email_id in email_id_list:
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
            msg2 = msg

            # Processando informações do email
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding)

            # Verificando o conteúdo do email
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    try:
                        if "attachment" not in content_disposition:
                            body = part.get_payload(decode=True).decode()
                            soup = BeautifulSoup(part.get_payload(decode=True).decode(), "html.parser")
                            print("Subject:", subject)
                            print("Content Type:", content_type)
                            print("Body:", body)
                            soup = BeautifulSoup(body, 'html.parser')
                            urls = soup.find_all('a')
                            if len(urls)==0:
                                corpo=soup.string
                                conti = corpo.find("Temporary URL: ")
                                contf = corpo.find("URL will")
                                if conti != -1 and contf != -1:
                                    url = corpo[conti + len(conti)+14:contf]
                                    print(url)
                                    temporary_url = url
                                    print("Temporary URL:", temporary_url)

                                else:
                                    print("Não foi possível encontrar o conteúdo.")
                            else:
                                print(urls)
                                temporary_url = urls[0]['href']
                                print("Temporary URL:", temporary_url)

                    except Exception as e:
                        print(f"Erro ao processar o conteúdo do email: {e}")
            else:
                try:
                    body = msg.get_payload(decode=True).decode()

                    # print("Subject:", subject)
                    # print("Content Type:", msg.get_content_type())
                    # print("Body:", body)

                    soup = BeautifulSoup(body, 'html.parser')
                    urls = soup.find_all('a')

                    # print(urls)

                    temporary_url = urls[0]['href']
                    print("Temporary URL:", temporary_url)

                    # for url in urls:
                    #     if 'Temporary URL' in url.get_text():
                    #         temporary_url = url['href']
                    #         print("Temporary URL:", temporary_url)
                except Exception as e:
                    print(f"Erro ao processar o conteúdo do email: {e}")

    except imaplib.IMAP4.error as e:
        print(f"Erro ao se conectar ao servidor IMAP: {e}")

    finally:
        # Fechando a conexão com o servidor IMAP
        print("Processo concluído com sucesso")
        mail.logout()
        return temporary_url

def baixa_arquivo(url):

    file_name = 'ERRO'
    try:
        response = requests.get(url)

        if response.status_code == 200:
            file_name = url.split("/")[-1].split("?")[0] # Obtém o nome do arquivo da URL
            file_path = os.path.abspath(file_name)
            with open(file_name, "wb") as file:
                file.write(response.content)
                print("Arquivo baixado com sucesso:", file_name)
            
        else:
            print("Falha ao baixar o arquivo. Código de status:", response.status_code)

    except requests.exceptions.RequestException as e:
        print("Erro ao fazer a requisição:", e)
    except Exception as e:
        print("Erro inesperado:", e)
    return file_path

def descriptografar(path_gpg):

    try:
        # variavel_de_instalacao

        # Crie um objeto GPG
        # gpg = gnupg.GPG(gnupghome="C:/Users/pelai/AppData/Roaming/gnupg")
        gpg = gnupg.GPG()

        # Caminho para o arquivo .asc com o certificado
        cert_file_path = 'Certificados/Secreta.asc'

        # Importe o certificado
        with open(cert_file_path, 'rb') as f:
            cert_data = f.read()
            import_result = gpg.import_keys(cert_data)

        # nome do arquivo a ser salvo
        file_path = path_gpg[:-4]

        # Realize a descriptografia usando o certificado importado
        with open(path_gpg, 'rb') as f:
            status = gpg.decrypt_file(f, output=file_path, passphrase="bporto@a5solutions.com")

        # Verifique o resultado da descriptografia
        if status.ok:
            print(f'Arquivo descriptografado salvo em {file_path}')
            return file_path
        else:
            print('Falha na descriptografia:', status.status)
            return 'ERRO'
    except Exception as e:
        print("Erro ao descriptografar arquivo - ", e)
        return 'ERRO'

def descompactar(zip_path):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(os.getcwd())
            nome_csv = zip_path.replace(".zip",".csv")

        try:
            os.path.exists(nome_csv)
            path_csv = os.path.abspath(nome_csv)
            return path_csv
        
        except Exception as a:
            print(f"Erro ao pegar o caminho do arquivo: {a}")
            return 'ERRO'
            
    except Exception as e:
        print(f"Erro ao descompactar arquivo: {e}")
        return 'ERRO'

def conecta_bd():
    try:
        with open("banco.json") as config_file:
            config = json.load(config_file)

        endereco    = config["endereco"]
        banco       = config["banco"]
        usuario     = config["usuario"]
        senha       = config["senha"]
        porta       = config["porta"]
        
        conexao = psycopg2.connect(
            dbname=banco,
            user=usuario,
            password=senha,
            host=endereco,
            port=porta
        )
        return conexao,usuario,banco,senha,endereco,int(porta)
    except FileNotFoundError:
        print("Arquivo banco.json não encontrado.")
        log.warning(f'Arquivo banco.json não encontrado.')
        return 'ERRO', 'ERRO', 'ERRO', 'ERRO'
    except json.JSONDecodeError:
        print("Erro ao decodificar o arquivo banco.json.")
        log.warning(f'Erro ao decodificar o arquivo banco.json.')
        return 'ERRO', 'ERRO', 'ERRO', 'ERRO'
    # except KeyError as e:
    #     print("Chave ausente no arquivo config.json:", e)
    #     return 'ERRO', 'ERRO', 'ERRO', 'ERRO'
    except Exception as e:
        print("Erro inesperado:", e)
        log.warning(f'Erro inesperado:', e)
        return 'ERRO', 'ERRO', 'ERRO', 'ERRO'

def inserir_bd(sql):
    con,usuario,banco,senha,endereco,porta = conecta_bd()
    cur = con.cursor()
    try:
        res = cur.execute(sql)
        con.commit()
        return res
    except (Exception, psycopg2.DatabaseError) as error:
        print("Error: %s" % error)
        log.warning(f'Error ao inserir ', error)       
        con.rollback()
        cur.close()
        return 1
    cur.close()

def editar_bd(sql):
    con,usuario,banco,senha,endereco,porta = conecta_bd()
    cur = con.cursor()
    try:
        cur.execute(sql)
        con.commit()
    except (Exception, psycopg2.DatabaseError) as error:
        print("Error: %s" % error)
        log.warning(f'Error no update ', error)       
        con.rollback()
        cur.close()
        return 1
    cur.close()

def seleciona_bd(sql):

    # conexao com o bd
    con,usuario,banco,senha,endereco,porta = conecta_bd()
    cur = con.cursor()
    
    # executa a query no bd
    try:
        cur.execute(sql)
        result = cur.fetchall()
    except psycopg2.Error as e:
        print(f"Erro ao executar a consulta SQL: {e}")

    cur.close()

    return result

def bkp_historico_bd(procedure):
    try:
        # Conecte-se ao banco de dados PostgreSQL
        conn ,usuario,banco,senha,endereco,porta = conecta_bd()

        # Crie um cursor
        cursor = conn.cursor()

        # Execute a função
        cursor.callproc(procedure)

        # Faça um commit para salvar as alterações no banco de dados
        conn.commit()

        # Feche o cursor e a conexão
        cursor.close()
        conn.close()
        print(f'Procedure "{procedure}" feita com sucesso!')
        return "OK"
    
    except Exception as e:
        print(f"Erro ao rodar procedure '{procedure}': {e}")
        return "ERRO"

def leitura_csv(nome_arquivo):
    df              = pd.read_csv(nome_arquivo)
    qtd_registros   = df.shape[0]
    df              = df.sort_values(by=['conversationId', 'eventId'])
    try:
        conn,usuario,banco,senha,endereco,porta = conecta_bd()
        cursor                                  = conn.cursor()
        cursor.execute("truncate table stg_diario")
        conn.commit()
        engine                                  = create_engine('postgresql://'+usuario+':'+senha+'@'+endereco+':'+str(porta)+'/'+banco)
        try:
            print('Aquivo',nome_arquivo)
            print('Inicio carga arquivo ')
            print('Quantidade de registros',qtd_registros)

            log.warning(f'Aquivo {nome_arquivo}')
            log.warning(f'Inicio carga arquivo')
            log.warning(f'Quantidade de registros: {str(qtd_registros)}')

            df.to_sql('stg_diario', engine, if_exists='replace', index=False)
            engine.dispose()
            log.warning(f'Fim carga arquivo')
            print('Fim carga arquivo ')

            ##Tabela ER bruta
            print('Iniciando carga tabela bruta')
            log.warning(f'Iniciando carga tabela bruta')

            # Invoca a função "inserir_dados" que está dentro do banco de dados: serve para levar os dados da stg_diario >> historico
            bkp_historico_bd("public.inserir_dados")

            log.warning(f'Fim carga tabela bruta')
            print('Fim carga tabela bruta')

        except Exception as d:
            print('Erro ao gravar no banco, realizando outra tentativa',d)
            log.warning(f'Erro ao gravar no banco, realizando outra tentativa {s}')       
            log.warning(f'Inicio 2 carga arquivo')
            print('Inicio 2 carga arquivo')
            try:
                for _, row in df.iterrows():
                    cursor.execute(
                    "INSERT INTO stg_diario (index,conversationId,conversationEndTime,conversationEndTimeL,dialogId,eventKey,eventId,time,timeL,participantId,skillId,skillName,messageId,sentBy,eventBy,permission,role,msgId,seq,agentId,agentFullName,agentGroupId,agentGroupName,agentLoginName,agentNickname,agentPid,assignedAgentFullName,assignedAgentId,assignedAgentLoginName,assignedAgentNickname,audience,avatarURL,consumerName,device,email,firstName,interactionTime,interactionTimeL,interactiveSequence,lastName,mcs,messageRawScore,messageTime,phone,quickReplies,rawMetadata,reason,responseTime,responseTimeAssignment,richContent,source,sourceAgentFullName,sourceAgentId,sourceAgentLoginName,sourceAgentNickname,sourceSkillId,sourceSkillName,targetSkillId,targetSkillName,token,words,questions,text,"'"intents-selectedClassification"'","'"intents-intentName"'","'"intents-intentLabel"'","'"intents-metaIntentName"'","'"intents-confidenceScore"'","'"intents-modelName"'","'"intents-modelVersion"'","'"messageStatus-messageDeliveryStatus"'","'"messageStatus-participantId"'","'"messageStatus-participantType"'","'"messageStatus-seq"'","'"messageStatus-time"'","'"messageStatus-timeL"'") VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",                (row['index'],row['conversationId'],row['conversationEndTime'],row['conversationEndTimeL'],row['dialogId'],row['eventKey'],row['eventId'],row['time'],row['timeL'],row['participantId'],row['skillId'],row['skillName'],row['messageId'],row['sentBy'],row['eventBy'],row['permission'],row['role'],row['msgId'],row['seq'],row['agentId'],row['agentFullName'],row['agentGroupId'],row['agentGroupName'],row['agentLoginName'],row['agentNickname'],row['agentPid'],row['assignedAgentFullName'],row['assignedAgentId'],row['assignedAgentLoginName'],row['assignedAgentNickname'],row['audience'],row['avatarURL'],row['consumerName'],row['device'],row['email'],row['firstName'],row['interactionTime'],row['interactionTimeL'],row['interactiveSequence'],row['lastName'],row['mcs'],row['messageRawScore'],row['messageTime'],row['phone'],row['quickReplies'],row['rawMetadata'],row['reason'],row['responseTime'],row['responseTimeAssignment'],row['richContent'],row['source'],row['sourceAgentFullName'],row['sourceAgentId'],row['sourceAgentLoginName'],row['sourceAgentNickname'],row['sourceSkillId'],row['sourceSkillName'],row['targetSkillId'],row['targetSkillName'],row['token'],row['words'],row['questions'],row['text'],row['intents-selectedClassification'],row['intents-intentName'],row['intents-intentLabel'],row['intents-metaIntentName'],row['intents-confidenceScore'],row['intents-modelName'],row['intents-modelVersion'],row['messageStatus-messageDeliveryStatus'],row['messageStatus-participantId'],row['messageStatus-participantType'],row['messageStatus-seq'],row['messageStatus-time'],row['messageStatus-timeL']))
                conn.commit()
                cursor.close()
                conn.close()
                log.warning(f'Fim 2 carga arquivo')
                print('Fim 2 carga arquivo')

                #Tabela stg_diario_historico
                print('Iniciando carga tabela bruta')
                log.warning(f'Iniciando carga tabela bruta')
                cursor.callproc('public.insere_dados')
                log.warning(f'Fim carga tabela bruta')
                print('Fim carga tabela bruta')

            except Exception as f:
                print('Erro ao gravar no banco ',f)
                log.warning(f'Erro ao gravar no banco {f}')
                return 'ERRO'
        conn.close()
    except Exception as e:
        print('Erro conexao com banco >>> ',e)
        log.warning(f'Erro conexao com banco >>> {e}')
        return 'ERRO'
    # return 'OK'

def processaplanilha():
    try:
        conn,usuario,banco,senha,endereco,porta=conecta_bd()
        cursor              = conn.cursor()
        cursor.execute('truncate table er_mov')
        conn.commit()
        cursor.execute('select "conversationId","conversationEndTime","eventId","eventKey","eventBy","skillId","skillName","agentId","agentFullName","sentBy","agentGroupId","agentGroupName","agentLoginName","agentNickname","time","permission","reason" from stg_diario order by "conversationId","eventId"')
        linhas              = cursor.fetchall()
        fgachei             = False
        fgbot               = False
        conversa_anterior   = 0
        contador            = 0
        contagent_humam     = 0
        permission          = 0
        dthpermission       = 0
        dthskill            = 0
        conversa_skill      = 0
        for linha in linhas:
            # cola
            # 0-conversationId, 1-conversationEndTime, 2-eventId, 3-eventKey, 4-eventBy
            # 5-skillId, 6-skillName, 7-agentId, 8-agentFullName, 9-sentBy, 10-agentGroupId
            # 11-agentGroupName, 12-agentLoginName, 13-agentNickname, 14-time, 15-permission
            # 16-reason
            contador    = contador+1
            conversa    = linha[0]
            if conversa_anterior==0:
                conversa_anterior=linha[0]
            if conversa==conversa_anterior and fgachei==False:
                if linha[4] == 'AgentBot':
                    dataehora=datetime.strptime(linha[14], '%Y-%m-%d %H:%M:%S.%f%z')
                    fgbot=True
                if linha[4] == 'AgentHuman' and fgbot:
                    contagent_humam=contagent_humam+1
                    dataehorahumano=datetime.strptime(linha[14], '%Y-%m-%d %H:%M:%S.%f%z')
                    tempo_transferencia_segundos= ((dataehorahumano-dataehora).total_seconds())/60
                    tempo_transferencia_minutos= dataehorahumano-dataehora
                    fgachei=True
                    #***Formata Transferencia Minutos
                    # quebra = str(tempo_transferencia_minutos).split(':')
                    # horas = int(quebra[0])
                    # minutos = int(quebra[1])
                    # segundos = int(quebra[2].split('.')[0])
                    # delta = timedelta(hours=horas, minutes=minutos, seconds=segundos)
                    # hora_formatada = str(delta).split('.')[0]
                    # #***
                    # if delta <= timedelta(hours=0, minutes=0, seconds=59):
                    #     faixa='0s ~ 60s'
                    # elif delta <= timedelta(hours=0, minutes=5, seconds=0):
                    #     faixa='60s ~ 5min'
                    # elif delta <= timedelta(hours=0, minutes=30, seconds=0):
                    #     faixa='5min ~ 30min'
                    # elif delta <= timedelta(hours=0, minutes=59, seconds=0):
                    #     faixa='30min ~ 60min'
                    # elif delta <= timedelta(hours=5, minutes=0, seconds=0):
                    #     faixa='60min ~ 5h'
                    # elif delta <= timedelta(hours=10, minutes=0, seconds=0):
                    #     faixa='5h ~ 10h'
                    # else:
                    #     faixa='maior que 10h'
                    if  permission != 0 and permission=='ASSIGNED_AGENT' and permission is not None:
                        dthpermission=datetime.strptime(dthpermission, '%Y-%m-%d %H:%M:%S.%f%z')
                        tempo_transferencia_segundos_permission= ((dataehorahumano-dthpermission).total_seconds())/60
                        tempo_transferencia_minutos_permission= dataehorahumano-dthpermission
                        if tempo_transferencia_minutos_permission.days >0:
                            horas_permission = int('24')
                            minutos_permission = int('59')
                            segundos_permission = int('59')
                        else:
                            quebra_permission = str(tempo_transferencia_minutos_permission).split(':')
                            horas_permission = int(quebra_permission[0])
                            minutos_permission = int(quebra_permission[1])
                            segundos_permission = int(quebra_permission[2].split('.')[0])
                        
                        delta_permission = timedelta(hours=horas_permission, minutes=minutos_permission, seconds=segundos_permission)
                        hora_formatada_permission= str(delta_permission).split('.')[0]                
                    if skill=='Skill' and linha[0]==conversa_skill:

                        if isinstance(dthskill,datetime):
                            print(str(dthskill))
                            dthskill=datetime.strptime(str(dthskill), '%Y-%m-%d %H:%M:%S.%f%z')
                        else:
                            dthskill=datetime.strptime(dthskill, '%Y-%m-%d %H:%M:%S.%f%z')
                        tempo_transferencia_segundos_skill= ((dataehorahumano-dthskill).total_seconds())/60
                        tempo_transferencia_minutos_skil= dataehorahumano-dthskill
                        
                        if tempo_transferencia_minutos_skil.days >0:# and contador==5741:
                            horas_skill = int('24')
                            minutos_skill = int('59')
                            segundos_skill = int('59')
                            horas = int('24')
                            minutos = int('59')
                            segundos = int('59')
                        else:
                            quebra_skill = str(tempo_transferencia_minutos_skil).split(':')
                            horas_skill = int(quebra_skill[0])
                            minutos_skill = int(quebra_skill[1])
                            segundos_skill = int(quebra_skill[2].split('.')[0])
                            quebra = str(tempo_transferencia_minutos_skil).split(':')
                            horas = int(quebra[0])
                            minutos = int(quebra[1])
                            segundos = int(quebra[2].split('.')[0])

                        delta_skill = timedelta(hours=horas_skill, minutes=minutos_skill, seconds=segundos_skill)
                        hora_formatada_skill= str(delta_skill).split('.')[0]                
                        #***Formata Transferencia Minutos fila
                        delta = timedelta(hours=horas, minutes=minutos, seconds=segundos)
                        hora_formatada = str(delta).split('.')[0]
                        #***
                        if delta <= timedelta(hours=0, minutes=0, seconds=59):
                            faixa='0s ~ 60s'
                        elif delta <= timedelta(hours=0, minutes=5, seconds=0):
                            faixa='60s ~ 5min'
                        elif delta <= timedelta(hours=0, minutes=30, seconds=0):
                            faixa='5min ~ 30min'
                        elif delta <= timedelta(hours=0, minutes=59, seconds=0):
                            faixa='30min ~ 60min'
                        elif delta <= timedelta(hours=5, minutes=0, seconds=0):
                            faixa='60min ~ 5h'
                        elif delta <= timedelta(hours=10, minutes=0, seconds=0):
                            faixa='5h ~ 10h'
                        else:
                            faixa='maior que 10h'
                    else:
                        hora_formatada_skill = "00:00"

                    sql=f'''insert into er_mov (conversationid,conversationendtime,eventkey,eventid,skillid,skillname,sentby,eventby,agentid,agentfullname,agentgroupid,agentgroupname,agentloginname,agentnickname,tmp_1a_resposta,dthinsercao,faixa,tmp_1a_resp_associado,tmp_fila) values('{linha[0]}','{linha[1]}','{linha[3]}','{linha[2]}','{linha[5]}','{linha[6]}','{linha[9]}','{linha[4]}','{str(int(linha[7]))}','{linha[8]}','{str(int(linha[10]))}','{linha[11]}','{linha[12]}','{linha[13]}','{hora_formatada}',now(),'{faixa}','{hora_formatada_permission}','{hora_formatada_skill}');'''
                    insere = inserir_bd(sql)
                if linha[15]=='ASSIGNED_AGENT':
                    permission=linha[15]
                    dthpermission=linha[14]
                if linha[16]=='Skill':
                    dthskill=linha[14]
                    skill = linha[16]
                    conversa_skill=linha[0]
            else:
                dataehora=0
                dataehorahumano=0
                conversa_anterior=linha[0]
                fgachei=False
                fgbot=False
                permission=0
            if contador % 1000 == 0:
                print('Registro',contador)
                log.warning(f'Registro: {contador}, Human: {contagent_humam}')

        # faz bkp dos dados da tabela er_mov para er_historico (após o truncate diario)
        bkp_historico_bd('public.inserir_dados_er_historico')

    except Exception as e:
        log.warning(f'Registro: {contador} - {e}')
        return 'ERRO'
    
    # return "OK"

def postgresql_to_xlsx():
    try:
        sql="select * from er_mov"

        # query no bd
        query_data = seleciona_bd(sql)

        # Transformar a lista de tuplas em uma lista de strings
        headers = ['conversationid', 'conversationendtime', 'eventkey', 'eventid', 'skillid', 'skillname', 'sentby', 'eventby', 'agentid', 'agentfullname', 'agentgroupid', 'agentgroupname', 'agentloginname', 'agentnickname', 'tmp_1a_resposta', 'tmp_1a_resp_associado', 'tmp_fila', 'dthinsercao', 'faixa']

        # Data Frame baseado na Query SQL
        df = pd.DataFrame(query_data) # insere os dados pegos do BD no DF
        df.columns = headers # define os headers do DF
        df = df.drop('dthinsercao', axis=1) # Remove a coluna 'dthinsercao'

        # nome do arquivo a ser salvo
        nome_arquivo = f"extracao_{data_hora_atual}.xlsx"

        # tranforma o df em um arquivo xlsx
        df.to_excel(nome_arquivo, index=False, sheet_name="Planilha1")
        
        # pega o caminho absoluto do arquivo xlsx
        file_path = os.path.abspath(nome_arquivo)

        print(f'Arquivo xlsx gerado com sucesso: {nome_arquivo}')
        log.warning(f"Dados puxados do BD e gerado o xlsx sem senha")
        # return file_path
        return file_path
    
    except Exception as e:
        print(f"Erro ao puxar os dados do BD e transformar em xlsx: {e}")
        log.warning(f"Erro ao puxar dados do BD e gerar o xlsx: {e}")
        return 'ERRO', 'ERRO'

def encrypt_xlsx(path, password):
    try:
        xl = w32.gencache.EnsureDispatch('Excel.Application')
        path = path if os.path.isabs(path) else os.path.abspath(path)
        
        wb = xl.Workbooks.Open(path)
        xl.DisplayAlerts = False
        wb.SaveAs(path, 51, password)
        xl.Quit()

        file_path = os.path.basename(path)
        
        print(f"Arquivo xlsx encriptografado com sucesso")
        log.warning(f"Arquivo xlsx encriptografado com sucesso")
        return file_path
    
    except Exception as e:
        print(f"Erro ao encriptografar arquivo xlsx: {e}")
        log.warning(f"Erro ao encriptografar arquivo xlsx: {e}")
        return "ERRO"

def compactar(xlsx_path):

    arquivo_zip = f"extracao_{data_hora_atual}.zip"

    try:
        # Crie o arquivo ZIP
        with zipfile.ZipFile(arquivo_zip, "w", zipfile.ZIP_DEFLATED) as arquivozip:
            # Adicione o arquivo XLSX ao arquivo ZIP com um nome específico dentro do ZIP
            arquivozip.write(xlsx_path, os.path.basename(xlsx_path))

        print(f"Arquivo compactado: {arquivo_zip}")
        log.warning(f"Arquivo xlsx compactado: {arquivo_zip}")
        return arquivo_zip

    except Exception as e:
        print(f"Erro ao compactar arquivo: {e}")
        log.warning(f"Erro ao compactar arquivo: {e}")
        return "ERRO"

def enviar_email(anexo, email_user, email_password, destinatarios):
    try:
        # variaveis
        assunto = "Extrator GetNet"

        # Configurações do servidor SMTP do Gmail
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587

        # Criar objeto do servidor SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)

        # Iniciar conexão com o servidor
        server.starttls()

        # Efetuar login na sua conta do Gmail
        server.login(email_user, email_password)

        # Criar um objeto MIMEMultipart
        msg = MIMEMultipart()

        # Configurar os campos do e-mail
        msg['From'] = email_user
        
        # tratamento de destinatarios
        emails = destinatarios.split(',')
        email_list = [email.strip() for email in emails]
        
        # Construir uma única string com todos os destinatários
        destinatarios_formatados = ', '.join(email_list)
        
        msg['To'] = destinatarios_formatados  # Defina a lista de destinatários aqui
        msg['Subject'] = assunto

        # Anexar um arquivo ao e-mail
        filename = anexo.split("\\")[-1]  # Extrai o nome do arquivo do caminho
        with open(anexo, 'rb') as attachment:
            part = MIMEApplication(attachment.read(), Name=filename)

        part['Content-Disposition'] = f'attachment; filename={filename}'
        msg.attach(part)

        # Enviar o e-mail
        server.sendmail(email_user, email_list, msg.as_string())

        # Fechar a conexão com o servidor
        server.quit()

        print('E-mails enviados com sucesso!')
        log.warning(f'E-mails enviados com sucesso. Emails: {email_list}')
        return email_list

    except Exception as e:
        print(f"Erro ao enviar xlsx via email: {e}")
        log.warning(f"Erro ao enviar xlsx via email: {e}")
        return 'ERRO'

def telegram_bot(email_list):
    try:    
        token = '6586450844:AAEIhkpEEWUlXZ4HTm0z55FaHTp14gMqfUc'
        url_base = f'https://api.telegram.org/bot{token}/'
        chat_id = '-4081934483'

        # msg aqui
        agora = datetime.now()
        data = agora.date().strftime('%d/%m/%y')
        horario = agora.time().strftime('%H:%M:%S')
        emails = '\n        - '.join([email for email in email_list])

        msg = f'''      ENVIO DE ARQUIVO - GETNET  

Infomações:
    >> Data: {data}
    >> Horário: {horario}
    >> Destinatários:
        - {emails}
'''

        link_envio = f'{url_base}sendMessage?chat_id={chat_id}&text={msg}'
        requests.get(link_envio)

        print(f'Mensagem enviada pro Telegram com sucesso!')
        log.warning(f'Mensagem enviada pro Telegram com sucesso!')
        return('OK')
    
    except Exception as e:
        print(f'Erro ao enviar mensagem pro telegram: {e}')
        log.warning(f'Erro ao enviar mensagem pro telegram: {e}')
        return 'ERRO'

def mk_historicoZip_file():
    '''
    Função para criar a pasta "historico_zip":
        - caso não exista: cria a pasta
        - caso ela exista: não faz nada
    '''
    try:
        # Diretório em que deseja verificar/criar a pasta
        diretorio_atual = os.getcwd()

        # Nome da pasta que você deseja verificar/criar
        nome_pasta = 'historico_zip'

        # Verificar se a pasta já existe
        if not os.path.exists(os.path.join(diretorio_atual, nome_pasta)):

            # Se a pasta não existe, crie-a
            os.makedirs(os.path.join(diretorio_atual, nome_pasta))
            historicoZip_path = f"{diretorio_atual}\{nome_pasta}"

            print(f'A pasta foi criada em {diretorio_atual}\{nome_pasta}')
            log.warning(f'A pasta foi criada em {diretorio_atual}\{nome_pasta}')
            return historicoZip_path
        
        else:
            historicoZip_path = f"{diretorio_atual}\{nome_pasta}"
            print(f'A pasta "{nome_pasta}" já existe em {diretorio_atual}')
            log.warning(f'A pasta "{nome_pasta}" já existe em {diretorio_atual}')
            return historicoZip_path
        
    except Exception as e:
        print(f"Erro ao criar pasta: {e}")
        return "ERRO"

def move_zipfile(arquivo_zipado_path, pasta_historico_path):
    try:
        final_path = f"{pasta_historico_path}\{os.path.basename(arquivo_zipado_path)[:-4]}_{data_hora_atual}.zip"

        # movimenta o arquivo de um path para outro
        final_path = shutil.move(arquivo_zipado_path, final_path)

        print(f"Arquivo zipada movida para historico_zip")
        log.warning(f"Arquivo zipada movida para historico_zip")
        return final_path
    
    except Exception as e:
        print(f"Erro ao movimentar o arquivo zipada: {e}")
        log.warning(f"Erro ao movimentar o arquivo zipada: {e}")
        return "ERRO"

def rmv_path(path):
    if os.path.exists(path):
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
            print(f"{path} removido com sucesso")
            log.warning(f"{path} removido com sucesso")

        except Exception as e:
            print(f"Erro ao remover {path}: {e}")
            log.warning(f"Erro ao remover {path}: {e}")
    else:
        print(f"O caminho {path} não existe")

def remover_arquivos(path_gpg, path_csv, path_xlsx, path_xlsx_compactado):
    try:
        rmv_path(path_gpg)
        rmv_path(path_csv)
        rmv_path(path_xlsx)
        rmv_path(path_xlsx_compactado)

        print("Arquivos removidas com sucesso!")
        log.warning("Arquivos removidas com sucesso!")
        return "OK"
    
    except Exception as e:
        print(f"Erro ao apagar os arquivos: {e}")
        log.warning(f"Erro ao apagar os arquivos: {e}")
        return "ERRO"

