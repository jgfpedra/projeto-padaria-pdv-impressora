import socket
import tkinter as tk
import time
import re
from collections import deque
from typing import Deque, Union, Sequence, Optional
import platform

PORT = 9100
IMPRESSORA_ATUAL: str = "epson"  # "epson" ou "bematech"

PARTES_CF_SIMPLIFICADO = [
    "DOCUMENTO AUXILIAR DA NOTA FISCAL",
    "Via Consumidor",
    "CHAVE DE ACESSO",
    "DESCRICAO",
    "Protocolo de Autorizacao",
    "VALOR TOTAL R$",
    "Consulta via leitor de QR Code"
]
PARTES_CF_CARTAO = PARTES_CF_SIMPLIFICADO + ["COMPROVANTE DE CREDITO OU DEBITO"]
PARTES_CF_1A_VIA = PARTES_CF_SIMPLIFICADO + ["1a VIA"]
PARTES_CF_2A_VIA = PARTES_CF_SIMPLIFICADO + [["VIA ESTABELECIMENTO", "2a VIA-ESTABELECIMENTO"]]
PARTES_POSSIVEIS: Sequence[Sequence[Union[str, list[str]]]] = [
    PARTES_CF_SIMPLIFICADO, PARTES_CF_CARTAO, PARTES_CF_1A_VIA, PARTES_CF_2A_VIA
]

dados_acumulados = b''
valid_jobs: Deque[bytes] = deque(maxlen=3)

def limpar_caracteres_de_controle(data: bytes) -> str:
    return re.sub(r'[\x00-\x1F\x7F\x1b]', '', data.decode('utf-8', errors='ignore'))

def valida_sequencia_bruta(data: bytes) -> bool:
    data_limpa = limpar_caracteres_de_controle(data)
    data_normalizada = " ".join(data_limpa.split())
    for partes_necessarias in PARTES_POSSIVEIS:
        partes_ok = True
        for parte in partes_necessarias:
            if isinstance(parte, list):
                if not any(alt in data_normalizada for alt in parte):
                    partes_ok = False
                    break
            else:
                if parte not in data_normalizada:
                    partes_ok = False
                    break
        if partes_ok:
            print("[DEBUG] Sequência completa encontrada para um dos tipos de cupom!")
            return True
    print("[DEBUG] Sequência incompleta para todos os tipos de cupom.")
    return False

def ask_receipt():
    result = None
    def yes(event: Optional[tk.Event] = None):
        nonlocal result
        result = True
        win.destroy()
    def no(event: Optional[tk.Event] = None):
        nonlocal result
        result = False
        win.destroy()
    win = tk.Tk()
    win.title("Cupom Fiscal")
    win.overrideredirect(True)
    width, height = 300, 150
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")
    win.attributes('-topmost', True)  # type: ignore[attr-defined]
    win.lift()  # type: ignore
    win.focus_force()
    win.grab_set()
    hidden_entry = tk.Entry(win)
    hidden_entry.place(x=-100, y=-100)
    hidden_entry.focus_set()
    if platform.system() == "Windows":
        try:
            import ctypes  # type: ignore
            hwnd = int(win.winfo_id())
            ctypes.windll.user32.SetForegroundWindow(hwnd)  # type: ignore[attr-defined]
        except Exception as e:
            print(f"[DEBUG] Não foi possível forçar foco absoluto: {e}")
    win.after(100, lambda: hidden_entry.focus_set())
    tk.Label(win, text="Cliente deseja o cupom?", font=("Arial", 14)).pack(pady=20)
    btns = tk.Frame(win)
    btns.pack()
    tk.Button(btns, text="Sim (F12)", width=12, command=lambda: yes()).grid(row=0, column=0, padx=10)
    tk.Button(btns, text="Não (F11)", width=12, command=lambda: no()).grid(row=0, column=1, padx=10)
    win.bind('<F12>', yes)
    win.bind('<F11>', no)
    win.bind('<Return>', yes)
    win.bind('<Escape>', no)
    win.mainloop()
    return result

def print_to_epson(data: bytes, printer_name: str = r"Epson") -> None:
    if platform.system() == "Windows":
        try:
            import win32print
            hPrinter = win32print.OpenPrinter(printer_name)
            try:
                win32print.StartDocPrinter(hPrinter, 1, ("Cupom", "", "RAW"))
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, data)
                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)  # type: ignore
            finally:
                win32print.ClosePrinter(hPrinter)
            print(f"Impressão enviada para: {printer_name}")
        except Exception as e:
            print(f"[ERRO] Impressão Epson falhou: {e}")
    else:
        print(f"[DEBUG] (Linux) Would print to Epson: {printer_name}, {len(data)} bytes")

def print_to_bematech(data: bytes, com_port: str = "COM3") -> None:
    if platform.system() == "Windows":
        try:
            import serial
            with serial.Serial(port=com_port, baudrate=9600, timeout=1) as ser:
                ser.write(data)
                ser.flush()
                time.sleep(0.5)
            print("Impressão enviada para Bematech via COM3")
        except Exception as e:
            print(f"[ERRO] Impressão Bematech falhou: {e}")
    else:
        print(f"[DEBUG] (Linux) Would print to Bematech: {com_port}, {len(data)} bytes")

def main():
    global dados_acumulados
    print(f"Middleware escutando na porta {PORT}...")
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(("0.0.0.0", PORT))
        s.listen(1)
        while True:
            try:
                conn, addr = s.accept()
                print(f"Recebendo job de {addr}")
                data = b''
                while True:
                    chunk = conn.recv(1024)
                    if not chunk:
                        break
                    data += chunk
                conn.close()
                if not data.strip():
                    print("[AVISO] Job recebido estava vazio. Ignorando.")
                    continue
                dados_acumulados += data
                print(f"[DEBUG] Job recebido (primeiros 200 bytes): {data[:200]}...")
                if valida_sequencia_bruta(dados_acumulados):
                    valid_jobs.append(dados_acumulados)
                    if len(valid_jobs) == 3:
                        wants = ask_receipt()
                        if wants:
                            for job in valid_jobs:
                                if IMPRESSORA_ATUAL == "bematech":
                                    print_to_bematech(job)
                                else:
                                    print_to_epson(job)
                        else:
                            print("Cliente recusou o cupom. Impressão abortada.")
                        valid_jobs.clear()
                    dados_acumulados = b''
                else:
                    print("[AVISO] Sequência inválida das partes. Imprimindo sem confirmação.")
                    if IMPRESSORA_ATUAL == "bematech":
                        print_to_bematech(dados_acumulados)
                    else:
                        print_to_epson(dados_acumulados)
                    dados_acumulados = b''
            except Exception as e:
                print(f"[ERRO] Falha na conexão ou processamento: {e}")
                continue

if __name__ == "__main__":
    main()
