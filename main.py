import funcoes as fun
from datetime import datetime

tempo_inicio = datetime.now()

fun.log.warning('Iniciando API')

try:
    email_user, email_password, mail_server, mail_port=fun.ler_config()
    if email_user != 'ERRO':
        link = fun.pega_link_email(email_user, email_password, mail_server, mail_port)
        print(link)
        print('---------------------------------------------------')
        if link != 'ERRO':
            path_gpg = fun.baixa_arquivo(link)
            print(path_gpg)
            print('---------------------------------------------------')
            if path_gpg != 'ERRO':
                path_zip =fun.descriptografar(path_gpg)
                print(path_zip)
                print('---------------------------------------------------')
                if path_zip != 'ERRO':
                    path_csv =fun.descompactar(path_zip)
                    print(path_csv)
                    print('---------------------------------------------------')
                    if path_csv != 'ERRO':
                        leitura_csv = fun.leitura_csv(path_zip)
                        print(leitura_csv)
                        print('---------------------------------------------------')
                        if leitura_csv != 'ERRO':
                            trata = fun.processaplanilha()
                            print(trata)
                            print('---------------------------------------------------')
                            if trata != 'ERRO':
                                path_xlsx = fun.postgresql_to_xlsx()
                                print(path_xlsx)
                                print('---------------------------------------------------')
                                if path_xlsx != 'ERRO':
                                    path_xlsx_compactado = fun.compactar(path_xlsx)
                                    print(path_xlsx_compactado)
                                    print('---------------------------------------------------')
                                    if path_xlsx_compactado != 'ERRO':
                                        email_list = fun.enviar_email(path_xlsx_compactado, email_user, email_password, destinatarios="gpelai@a5solutions.com")
                                        print(email_list)
                                        print('---------------------------------------------------')
                                        if email_list != 'ERRO':
                                            retorno_telegram = fun.telegram_bot(email_list)
                                            print(retorno_telegram)
                                            print('---------------------------------------------------')
                                            if retorno_telegram != 'ERRO':
                                                historicoZip_path = fun.mk_historicoZip_file()
                                                print(historicoZip_path)
                                                print('---------------------------------------------------')
                                                if historicoZip_path != 'ERRO':
                                                    movimento_zip = fun.move_zipfile(path_zip, historicoZip_path)
                                                    print(movimento_zip)
                                                    print('---------------------------------------------------')
                                                    if movimento_zip != "ERRO":    
                                                        remover_arquivos = fun.remover_arquivos(path_gpg, path_csv, path_xlsx, path_xlsx_compactado)
                                                        tempo_fim = datetime.now()
                                                        print('---------------------------------------------------')
                                                        tempo_total = tempo_fim - tempo_inicio
                                                        print(f"Tempo total do processo: {tempo_total.total_seconds()}")

except Exception as e:
    print(f'ERRO processo: {e}')