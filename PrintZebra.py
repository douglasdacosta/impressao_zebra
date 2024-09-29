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
            
            # Itera pelos itens retornados pela API
            for item in itens:
                material = item['material']
                fornecedor = item['fornecedor']
                estoque_id = item['estoque_id']
                data = item['data']

                # Endereço IP e porta da impressora Zebra
                printer_ip = config['CONFIG']['ipimpressora']
                printer_port = config['CONFIG']['porta']
                printer_name = config['CONFIG']['printer_name']
                # Comando ZPL para criar a etiqueta
                zpl = f"""
                ^XA
                ^PW300
                ^LL200
                ^FO0,50^FB300,2,0,C^A0N,30,30^FD{material}\&^FS
                ^FO0,90^FB300,2,0,C^A0N,20,20^FD{fornecedor}\&^FS
                ^FO0,120^FB300,2,0,C^A0N,30,20^FDID: {estoque_id} DATA: {data}\&^FS
                ^XZ
                """

                # Enviar a etiqueta para a impressora
                # send_zpl_to_printer(zpl, printer_ip, printer_port)
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