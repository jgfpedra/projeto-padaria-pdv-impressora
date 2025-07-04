import os
import socket
import threading
import serial
import time
from datetime import datetime

PORT = 9100  # Porta única para Epson ou Bematech
PRINTER_TYPE = "epson"  # Altere para "bematech" se necessário
printer_lock = threading.Lock()

def save_dump(data: bytes, directory: str = "dumps"):
    # Salva o dump na mesma pasta onde o script está
    base_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(base_dir, directory)
    
    os.makedirs(full_path, exist_ok=True)

    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S_%f")[:-3]
    filename = os.path.join(full_path, f"dump_{timestamp}.bin")
    
    with open(filename, "wb") as f:
        f.write(data)
    
    print(f"Dump salvo em {filename}")
    return filename

def send_to_printer(data: bytes, printer_type: str = PRINTER_TYPE):
    if printer_type == "epson":
        printer_path = r"\\127.0.0.1\Epson"  # Epson na rede
        try:
            with open(printer_path, "wb") as printer:
                printer.write(data)
            print(f"Dados enviados para a impressora Epson em {printer_path}")
        except Exception as e:
            print(f"Erro ao enviar para a impressora Epson: {e}")

    elif printer_type == "bematech":
        try:
            with printer_lock:
                # A Bematech está conectada via porta serial COM3
                with serial.Serial(port="COM3", baudrate=9600, timeout=1) as ser:
                    ser.write(data)
                    ser.flush()  # Força envio imediato do buffer
                    time.sleep(0.5)  # Dá tempo para a impressora processar
            print("Dados enviados para a Bematech MP-4200 em COM3")
        except Exception as e:
            print(f"Erro ao enviar para a Bematech: {e}")

def handle_connection(conn: socket.socket, addr: tuple[str, int]):
    data = b''  # Vai armazenar os dados recebidos
    try:
        print(f"Recebendo job de {addr}")
        conn.settimeout(5)  # Timeout de 5 segundos para leitura dos dados
        while True:
            chunk = conn.recv(1024)  # Lê os dados recebidos
            if not chunk:
                break
            data += chunk  # Concatena os chunks recebidos

    except Exception as e:
        print(f"Erro durante a conexão com {addr}: {e}")
    finally:
        conn.close()  # Fecha a conexão
        if data:
            try:
                save_dump(data)  # Salva os dados recebidos como dump
            except Exception as e:
                print(f"Erro ao salvar dump: {e}")
            try:
                send_to_printer(data)  # Envia os dados para a impressora correta
            except Exception as e:
                print(f"Erro ao enviar para a impressora: {e}")

def main():
    print(f"Middleware escutando na porta {PORT} para {PRINTER_TYPE}...")
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    s.bind(("0.0.0.0", PORT))
    s.listen(5)
    print(f"Escutando na porta {PORT}...")
    while True:
        conn, addr = s.accept()
        thread = threading.Thread(target=handle_connection, args=(conn, addr))
        thread.daemon = True
        thread.start()

if __name__ == "__main__":
    main()
