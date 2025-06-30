import os
import socket
import win32print
import tkinter as tk

PORT = 9100  # Porta local

CUT_CMD = b'\x1D\x56'  # Comando de corte ESC/POS

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
    win.geometry("280x120")
    win.attributes("-topmost", True)

    tk.Label(win, text="Cliente deseja o cupom?", font=("Arial", 12)).pack(pady=10)

    btns = tk.Frame(win)
    btns.pack()
    tk.Button(btns, text="Sim", width=10, command=yes).grid(row=0, column=0, padx=10)
    tk.Button(btns, text="Não", width=10, command=no).grid(row=0, column=1, padx=10)

    # Hotkeys
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

def print_to_real_printer(data, printer_name=r"\\MYPC\epson"):
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
            identified = identify_parts(parts)

            wants = ask_receipt()

            to_print = []
            if wants:
                print("Cliente deseja o cupom completo.")
                to_print = parts  # Imprime tudo
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
