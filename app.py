import socket
import win32print
import tkinter as tk
import serial
import time

PORT = 9100
CUT_CMD = b'\x1D\x56'  # comando de corte

# Escolha da impressora: "bematech" ou "epson"
IMPRESSORA_ATUAL = "bematech"  # altere aqui para "epson" se quiser
# Comandos ESC/POS básicos para Bematech

# Fonte normal (tamanho padrão)
FONT_NORMAL = b'\x1D\x21\x00'  

# Espaçamento entre linhas (exemplo: 30 pontos)
LINE_SPACING = b'\x1B\x33\x1E'  

# Espaçamento entre caracteres (exemplo: 2 pontos)
CHAR_SPACING = b'\x1B\x20\x02'  

def ajustar_fonte_espaco(data: bytes) -> bytes:
    # Comandos ESC/POS para Bematech:
    
    # ESC ! n  --> Define modo da fonte
    # n = 0x01 (fonte menor, modo normal)
    cmd_modo_fonte_pequena = b'\x1B\x21\x11' 
    
    # ESC SP n --> Define espaço entre caracteres
    # n = 0 (menor espaçamento)
    cmd_espaco_entre_chars = b'\x1B\x20\x00' 
    
    # ESC 3 n --> Define espaçamento entre linhas (altura da linha)
    # n = 15 (ajuste pequeno, padrão é 30)
    cmd_espaco_entre_linhas = b'\x1B\x33\x0F'
    
    # Monta os comandos no início do dado
    return cmd_modo_fonte_pequena + cmd_espaco_entre_chars + cmd_espaco_entre_linhas + data


def split_receipts(data):
    parts = data.split(CUT_CMD)
    return [part + CUT_CMD for part in parts if part.strip()]

def valida_sequencia(parts):
    texto_partes = [part.decode(errors='ignore') for part in parts]
    chave_nf = "DOCUMENTO AUXILIAR DA NOTA FISCAL"
    chave_tef = "COMPROVANTE CREDITO OU DEBITO"
    chave_via = "VIA ESTABELECIMENTO"

    try:
        idx_nf = next(i for i, texto in enumerate(texto_partes) if chave_nf in texto)
        idx_tef = next(i for i, texto in enumerate(texto_partes) if chave_tef in texto)
        idx_via = next(i for i, texto in enumerate(texto_partes) if chave_via in texto)
    except StopIteration:
        return False

    return idx_nf < idx_tef < idx_via

def ask_receipt():
    result = None

    def yes(event=None):
        nonlocal result
        result = True
        win.destroy()

    def no(event=None):
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
    win.attributes("-topmost", True)
    win.focus_force()

    tk.Label(win, text="Cliente deseja o cupom?", font=("Arial", 14)).pack(pady=20)

    btns = tk.Frame(win)
    btns.pack()
    tk.Button(btns, text="Sim (F12)", width=12, command=yes).grid(row=0, column=0, padx=10)
    tk.Button(btns, text="Não (F11)", width=12, command=no).grid(row=0, column=1, padx=10)

    win.bind('<F12>', yes)
    win.bind('<F11>', no)

    win.mainloop()
    return result

def print_to_epson(data, printer_name=r"Epson"):
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            win32print.StartDocPrinter(hPrinter, 1, ("Cupom", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, data)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)
        print(f"Impressão enviada para: {printer_name}")
    except Exception as e:
        print(f"[ERRO] Impressão Epson falhou: {e}")

def print_to_bematech(data, com_port="COM3"):
    import serial
    try:
        with serial.Serial(port=com_port, baudrate=9600, timeout=1) as ser:
            ser.write(data)
            ser.flush()
            time.sleep(0.5)
        print("Impressão enviada para Bematech via COM3")
    except Exception as e:
        print(f"[ERRO] Impressão Bematech falhou: {e}")

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

            parts = split_receipts(data)

            if valida_sequencia(parts):
                wants = ask_receipt()
                if not wants:
                    print("Cliente recusou o cupom. Impressão abortada.")
                    continue
            else:
                print("[AVISO] Sequência inválida das partes. Imprimindo sem confirmação.")

            final_data = b''.join(parts)

            if IMPRESSORA_ATUAL == "bematech":
                final_data = ajustar_fonte_espaco(final_data)  # Aplica comandos ESC/POS
                print_to_bematech(final_data)
            else:
                print_to_epson(final_data)

if __name__ == "__main__":
    main()