import os
import socket
import threading
import serial
import time
from datetime import datetime

PORT = 9100  # Altere se necessário
printer_lock = threading.Lock()

def save_dump(data, directory="dumps"):
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

def send_to_printer(data, printer_type="bematech"):
    if printer_type == "epson":
        printer_path = r"\\127.0.0.1\Epson"
        try:
            with open(printer_path, "wb") as printer:
                printer.write(data)
            print(f"Dados enviados para a impressora Epson em {printer_path}")
        except Exception as e:
            print(f"Erro ao enviar para a impressora Epson: {e}")

    elif printer_type == "bematech":
    	try:
        	with printer_lock:
            		with serial.Serial(port="COM3", baudrate=9600, timeout=1) as ser:
                		ser.write(data)
                		ser.flush()         # <- força envio imediato do buffer
                		time.sleep(0.5)     # <- dá tempo pra impressora processar
        	print("Dados enviados para a Bematech MP-4200 em COM3")
    	except Exception as e:
        	print(f"Erro ao enviar para a Bematech: {e}")

def handle_connection(conn, addr):
    try:
        print(f"Recebendo job de {addr}")
        data = b''
        conn.settimeout(5)
        while True:
            chunk = conn.recv(1024)
            if not chunk:
                break
            data += chunk
    except Exception as e:
        print(f"Erro durante a conexão com {addr}: {e}")
    finally:
        conn.close()
        if data:
            try:
                save_dump(data)
            except Exception as e:
                print(f"Erro ao salvar dump: {e}")
            try:
                send_to_printer(data)
            except Exception as e:
                print(f"Erro ao enviar para a impressora: {e}")

def main():
    print(f"Middleware escutando na porta {PORT}...")
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("0.0.0.0", PORT))
        s.listen(5)
        while True:
            conn, addr = s.accept()
            thread = threading.Thread(target=handle_connection, args=(conn, addr))
            thread.daemon = True
            thread.start()

if __name__ == "__main__":
    main()
