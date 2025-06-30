import os
import socket
import win32print
import tkinter as tk

PORT = 9100  # Porta local

CUT_CMD = b'\x1D\x56'  # Comando de corte ESC/POS

def valida_sequencia(parts):
    texto_partes = [part.decode(errors='ignore') for part in parts]

    # Palavras-chave para identificar cada parte
    chave_nf = "DOCUMENTO AUXILIAR DA NOTA FISCAL"
    chave_tef = "COMPROVANTE CREDITO OU DEBITO"
    chave_via = "VIA ESTABELECIMENTO"

    # Procura os índices dessas partes no texto das partes
    try:
        idx_nf = next(i for i, texto in enumerate(texto_partes) if chave_nf in texto)
        idx_tef = next(i for i, texto in enumerate(texto_partes) if chave_tef in texto)
        idx_via = next(i for i, texto in enumerate(texto_partes) if chave_via in texto)
    except StopIteration:
        # Se alguma parte não foi encontrada, retorna False
        return False

    # Verifica se a ordem é correta
    if idx_nf < idx_tef < idx_via:
        return True
    else:
        return False

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

    # Remove barra de título (sem close/min/max buttons)
    win.overrideredirect(True)

    # Define tamanho e posição centralizada
    width, height = 300, 150
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")

    # Mantém sempre no topo
    win.attributes("-topmost", True)

    # Garante foco na janela
    win.focus_force()

    tk.Label(win, text="Cliente deseja o cupom?", font=("Arial", 14)).pack(pady=20)

    btns = tk.Frame(win)
    btns.pack()

    tk.Button(btns, text="Sim (F12)", width=12, command=yes).grid(row=0, column=0, padx=10)
    tk.Button(btns, text="Não (F11)", width=12, command=no).grid(row=0, column=1, padx=10)

    # Hotkeys para teclado físico
    win.bind('<F12>', yes)  # F12 para "Sim"
    win.bind('<F11>', no)   # F11 para "Não"

    win.mainloop()
    return result

def split_receipts(data):
    parts = data.split(CUT_CMD)
    return [part + CUT_CMD for part in parts if part.strip()]

def identify_parts(parts):
    result = {
        "nfe": [],
        "tef": [],
        "via_estab": [],
        "outros": []
    }

    for part in parts:
        text = part.decode('latin1', errors='ignore')

        if "DOCUMENTO AUXILIAR DA NOTA FISCAL" in text:
            result["nfe"].append(part)
        elif "COMPROVANTE CREDITO OU DEBITO" in text:
            result["tef"].append(part)
        elif "VIA ESTABELECIMENTO" in text:
            result["via_estab"].append(part)
        else:
            result["outros"].append(part)

    return result

def print_to_real_printer(data, printer_name=r"Epson"):
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("Cupom", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, data)
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

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

            # Valida sequência das partes
            if not valida_sequencia(parts):
                print("[ERRO] Sequência inválida das partes. Impressão abortada.")
                continue  # Ignora esse job e volta para esperar o próximo

            identified = identify_parts(parts)
            wants = ask_receipt()

            if wants:
                print("Cliente deseja o cupom completo.")
                to_print = parts
            else:
                print("Cliente recusou. Imprimindo somente VIA ESTABELECIMENTO.")
                to_print = identified["via_estab"]

            if to_print:
                final_data = b''.join(to_print)
                print_to_real_printer(final_data)
            else:
                print("Nenhuma parte selecionada para impressão.")

if __name__ == "__main__":
    main()
