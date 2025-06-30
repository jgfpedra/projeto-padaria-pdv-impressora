import os
import socket
import threading
from datetime import datetime

PORT = 9100  # Altere se necessário

def save_dump(data):
    os.makedirs("dump", exist_ok=True)
    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S_%f")[:-3]
    filename = os.path.join("dump", f"dump_{timestamp}.bin")
    with open(filename, "wb") as f:
        f.write(data)
    print(f"Dump salvo em {filename}")
    return filename

def send_to_printer(data):
    printer_path = r"\\127.0.0.1\Epson"
    try:
        with open(printer_path, "wb") as printer:
            printer.write(data)
        print(f"Dados enviados para a impressora em {printer_path}")
    except Exception as e:
        print(f"Erro ao enviar para a impressora: {e}")

def handle_connection(conn, addr):
    print(f"Recebendo job de {addr}")
    data = b''
    while True:
        chunk = conn.recv(1024)
        if not chunk:
            break
        data += chunk
    conn.close()
    save_dump(data)
    send_to_printer(data)

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
