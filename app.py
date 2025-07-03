import socket
import tkinter as tk
import time
import re
from collections import deque  # Para armazenar os últimos 3 jobs
from typing import Deque
from typing import List
import platform
import sys
import os

PORT = 9100
CUT_CMD = b'\x1D\x56'  # Comando de corte
IMPRESSORA_ATUAL: str = "epson"  # Pode ser "epson" ou "bematech"
FONT_NORMAL = b'\x1D\x21\x00'
LINE_SPACING = b'\x1B\x33\x1E'
CHAR_SPACING = b'\x1B\x20\x02'

# Definição das partes obrigatórias para cada tipo de cupom
PARTES_CF_SIMPLIFICADO = [
    "DOCUMENTO AUXILIAR DA NOTA FISCAL",
    "Via Consumidor",
    "CHAVE DE ACESSO",
    "DESCRICAO",
    "Protocolo de Autorizacao",
    "VALOR TOTAL R$",
    "Consulta via leitor de QR Code"
]
PARTES_CF_CARTAO = PARTES_CF_SIMPLIFICADO + [
    "COMPROVANTE DE CREDITO OU DEBITO"
]
PARTES_CF_1A_VIA = PARTES_CF_SIMPLIFICADO + [
    "1a VIA"
]
PARTES_CF_2A_VIA = PARTES_CF_SIMPLIFICADO + [
    "VIA ESTABELECIMENTO"
]
PARTES_POSSIVEIS = [PARTES_CF_SIMPLIFICADO, PARTES_CF_CARTAO, PARTES_CF_1A_VIA, PARTES_CF_2A_VIA]

# Variáveis globais
# Não é mais necessário partes_encontradas nem sequencia_completa_encontrada
# dados_acumulados e valid_jobs permanecem

dados_acumulados = b''
valid_jobs: Deque[bytes] = deque(maxlen=3)  # Para armazenar os últimos 3 jobs válidos

# =====================
# Funções utilitárias
# =====================
def ajustar_fonte_espaco(data: bytes) -> bytes:
    cmd_modo_fonte_pequena = b'\x1B\x21\x11'
    cmd_espaco_entre_chars = b'\x1B\x20\x00'
    cmd_espaco_entre_linhas = b'\x1B\x33\x0F'
    return cmd_modo_fonte_pequena + cmd_espaco_entre_chars + cmd_espaco_entre_linhas + data

def limpar_caracteres_de_controle(data: bytes) -> str:
    return re.sub(r'[\x00-\x1F\x7F\x1b]', '', data.decode('utf-8', errors='ignore'))

def valida_sequencia_bruta(data: bytes) -> bool:
    data_limpa = limpar_caracteres_de_controle(data)
    data_normalizada = " ".join(data_limpa.split())
    for partes_necessarias in PARTES_POSSIVEIS:
        partes_ok = True
        for parte in partes_necessarias:
            if parte not in data_normalizada:
                partes_ok = False
                break
        if partes_ok:
            print("[DEBUG] Sequência completa encontrada para um dos tipos de cupom!")
            return True
    print("[DEBUG] Sequência incompleta para todos os tipos de cupom.")
    return False

# =====================
# Função de confirmação
# =====================
def ask_receipt():
    result = None
    from typing import Optional
    import platform
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
    win.attributes('-topmost', True) # type: ignore
    win.lift() # type: ignore
    win.focus_force()
    # Força o foco absoluto no Windows
    if platform.system() == "Windows":
        try:
            import ctypes
            hwnd = int(win.winfo_id())
            ctypes.windll.user32.SetForegroundWindow(hwnd)  # type: ignore[attr-defined]
        except Exception as e:
            print(f"[DEBUG] Não foi possível forçar foco absoluto: {e}")
    # Tenta forçar o foco após a janela ser exibida
    win.after(100, lambda: win.focus_force())
    tk.Label(win, text="Cliente deseja o cupom?", font=("Arial", 14)).pack(pady=20)
    btns = tk.Frame(win)
    btns.pack()
    tk.Button(btns, text="Sim (F12)", width=12, command=yes).grid(row=0, column=0, padx=10)
    tk.Button(btns, text="Não (F11)", width=12, command=no).grid(row=0, column=1, padx=10)
    win.bind('<F12>', yes)
    win.bind('<F11>', no)
    win.mainloop()
    return result

# =====================
# Funções de impressão
# =====================
if platform.system() == "Windows":
    import win32print
    import serial
    def print_to_epson(data: bytes, printer_name: str = r"Epson") -> None:
        try:
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
    def print_to_bematech(data: bytes, com_port: str = "COM3") -> None:
        try:
            with serial.Serial(port=com_port, baudrate=9600, timeout=1) as ser:
                ser.write(data)
                ser.flush()
                time.sleep(0.5)
            print("Impressão enviada para Bematech via COM3")
        except Exception as e:
            print(f"[ERRO] Impressão Bematech falhou: {e}")
else:
    def print_to_epson(data: bytes, printer_name: str = r"Epson") -> None:
        print(f"[DEBUG] (Linux) Would print to Epson: {printer_name}, {len(data)} bytes")
    def print_to_bematech(data: bytes, com_port: str = "COM3") -> None:
        print(f"[DEBUG] (Linux) Would print to Bematech: {com_port}, {len(data)} bytes")

# =====================
# Função principal do servidor
# =====================
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
                print(f"[DEBUG] Job completo: {dados_acumulados}")
                if valida_sequencia_bruta(dados_acumulados):
                    valid_jobs.append(dados_acumulados)
                    if len(valid_jobs) == 3:
                        wants = ask_receipt()
                        if wants:
                            for job in valid_jobs:
                                if IMPRESSORA_ATUAL == "bematech":
                                    final_data = ajustar_fonte_espaco(job)
                                    print_to_bematech(final_data)
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

# =====================
# Função de teste com arquivos dump
# =====================
def test_with_dump_files(dump_dir: str = "dump") -> None:
    global dados_acumulados
    files = sorted([f for f in os.listdir(dump_dir) if f.endswith(".bin")])
    if not files:
        print("[TEST] No .bin files found in dump directory.")
        return
    print(f"[TEST] Simulando jobs sequenciais: {len(files)} arquivos")
    buffer = b''
    bloco_jobs: List[str] = []
    algum_cupom_impresso = False
    for fname in files:
        with open(os.path.join(dump_dir, fname), "rb") as f:
            data = f.read()
        data_limpa = limpar_caracteres_de_controle(data)
        # Se encontrar o início de cupom, processa o bloco anterior (se houver) e inicia novo bloco a partir deste arquivo
        if PARTES_CF_SIMPLIFICADO[0] in data_limpa:
            if bloco_jobs:
                print(f"[TEST] Bloco detectado: arquivos {bloco_jobs}")
                if processa_bloco(buffer, bloco_jobs):
                    algum_cupom_impresso = True
            buffer = data
            bloco_jobs = [fname]
        else:
            buffer += data
            bloco_jobs.append(fname)
    # Processa o último bloco
    if bloco_jobs:
        print(f"[TEST] Bloco detectado: arquivos {bloco_jobs}")
        if processa_bloco(buffer, bloco_jobs):
            algum_cupom_impresso = True
    # Só imprime os últimos 3 arquivos se nenhum cupom foi impresso com confirmação
    # E só se houver mais de 7 arquivos (ou seja, não é um CF simples)
    if not algum_cupom_impresso and len(files) > 7:
        last3_files = files[-3:]
        last3_data = b''
        for fname in last3_files:
            with open(os.path.join(dump_dir, fname), "rb") as f:
                last3_data += f.read()
        print(f"[TEST] Printing concatenated last 3 files, total {len(last3_data)} bytes")
        if IMPRESSORA_ATUAL == "bematech":
            print_to_bematech(last3_data)
        else:
            print_to_epson(last3_data)
    dados_acumulados = b''

def processa_bloco(data: bytes, job_files: List[str]) -> bool:
    print(f"[TEST] Processando bloco de {len(job_files)} arquivos, {len(data)} bytes")
    if PARTES_CF_SIMPLIFICADO[0] in limpar_caracteres_de_controle(data):
        if valida_sequencia_bruta(data):
            print("[TEST] Sequência completa encontrada no bloco. Perguntando ao cliente...")
            wants = ask_receipt()
            if wants:
                if IMPRESSORA_ATUAL == "bematech":
                    final_data = ajustar_fonte_espaco(data)
                    print_to_bematech(final_data)
                else:
                    print_to_epson(data)
                return True  # Cupom impresso com confirmação
            else:
                print("[TEST] Cliente recusou o cupom. Impressão abortada.")
                return False
        else:
            print("[TEST] Sequência incompleta no bloco. Imprimindo mesmo assim...")
            if IMPRESSORA_ATUAL == "bematech":
                print_to_bematech(data)
            else:
                print_to_epson(data)
            return False
    else:
        print("[TEST] Bloco não contém início de cupom fiscal. Ignorando.")
        return False

# =====================
# Entry point
# =====================
if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test-dump":
        test_with_dump_files()
    else:
        main()
