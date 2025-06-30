import socket
from datetime import datetime

PORT = 9100  # Porta onde o middleware vai escutar

def save_dump(data):
    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S_%f")[:-3]  # até milissegundos
    filename = f"dump_{timestamp}.bin"
    with open(filename, "wb") as f:
        f.write(data)
    print(f"Dump salvo em {filename}")

def main():
    print(f"Middleware escutando na porta {PORT}...")
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("0.0.0.0", PORT))
        s.listen(1)
        while True:
            conn, addr = s.accept()
            print(f"Recebendo job de {addr}")
            data = b''
            while True:
                chunk = conn.recv(1024)
                if not chunk:
                    break
                data += chunk
            conn.close()
            save_dump(data)

if __name__ == "__main__":
    main()
