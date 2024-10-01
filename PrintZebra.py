import socket
import configparser
import requests
import json
import time
import logging
import win32print



# Configuração básica do logging
logging.basicConfig(level=logging.INFO)

def read_config(config_file):
    config = configparser.ConfigParser()
    config.read(config_file, 'utf-8')
    return config

def consultar_horas_turno(TOKEN, LOGIN, SENHA, url):
    parametros = {
        'TOKEN': TOKEN,
        'LOGIN': LOGIN,
        'SENHA': SENHA,
    }

    try:
        resposta = requests.get(url, params=parametros)
        if resposta.status_code == 200:
            return resposta.text
        else:
            resposta.raise_for_status()
    except requests.exceptions.RequestException as e:
        logging.error(f"Erro durante a solicitação: {e}")
        return None

# Enviando o comando ZPL para a impressora
def send_zpl_to_printer(zpl, ip, port):
    try:
        # Cria uma conexão com a impressora
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.connect((ip, int(port)))  # Garantindo que a porta é um número inteiro
            s.sendall(zpl.encode())
            logging.info("Etiqueta enviada para impressão.")
    except Exception as e:
        logging.error(f"Erro ao enviar a etiqueta: {e}")

def send_zpl_to_printer_windows(zpl, printer_name):
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            job = win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, zpl.encode())
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            logging.info("Etiqueta enviada para impressão no Windows.")
        finally:
            win32print.ClosePrinter(hPrinter)
    except Exception as e:
        logging.error(f"Erro ao enviar a etiqueta no Windows: {e}")

def main():
    try:
        # Carrega configurações
        config = read_config('config.conf')

        url = config['API']['urlimpressao']
        TOKEN = config['API']['TOKEN']
        LOGIN = config['API']['LOGIN']
        SENHA = config['API']['SENHA']

        resultado = consultar_horas_turno(TOKEN, LOGIN, SENHA, url)
        if resultado:
            itens = json.loads(resultado)
            
            # printer_ip = config['CONFIG']['ipimpressora']
            # printer_port = config['CONFIG']['porta']
            printer_name = config['CONFIG']['printer_name']
                
            # Itera pelos itens retornados pela API
            for item in itens:
                material = item['material']
                fornecedor = item['fornecedor']
                estoque_id = item['estoque_id']
                qtde_etiqueta = item['qtde']
                data = item['data']
                if qtde_etiqueta % 2 != 0:
                    qtde_etiqueta = qtde_etiqueta+1

                qtde_etiqueta = qtde_etiqueta//2

                print('qtde_etiqueta')    
                print(qtde_etiqueta)
                #passa pela quantidade de etiqueta enviando o comando para a impressora
                for i in range(qtde_etiqueta):
                    # Comando ZPL para criar a etiqueta
                    zpl = f"""
^XA
^PW800
^LL100
^FO0,20^FB400,1,0,C^A0N,30,30^FD{material}\&^FS
^FO0,55^FB400,2,0,C^A0N,20,20^FD{fornecedor}\&^FS
^FO0,80^FB400,2,0,C^A0N,30,20^FDLOTE: {estoque_id} {data}\&^FS
^FO223,20^FB650,2,0,C^A0N,30,30^FD{material}\&^FS
^FO223,55^FB650,2,0,C^A0N,20,20^FD{fornecedor}\&^FS
^FO223,80^FB650,2,0,C^A0N,30,20^FDLOTE: {estoque_id} {data}\&^FS
^XZ
"""                  
                    
                    # Enviar a etiqueta para a impressora                    
                    send_zpl_to_printer_windows(zpl, printer_name)

                logging.info(f"Etiqueta para {material} enviada.")
        
        else:
            logging.warning("Nenhuma etiqueta para processar.")
        
    except configparser.Error as e:
        logging.error(f"Erro ao ler o arquivo de configuração: {e}")
    except Exception as e:
        logging.error(f"Erro inesperado: {e}")

if __name__ == "__main__":
    while True:
        main()          # Executa a função main
        time.sleep(5)   # Aguarda 5 segundos antes da próxima execução