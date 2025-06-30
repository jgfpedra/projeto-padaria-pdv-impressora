import socket
import win32print
import tkinter as tk

PORT = 9100  # Porta local

def ask_receipt():
    result = None
    def yes(): nonlocal result; result = True; win.destroy()
    def no():  nonlocal result; result = False; win.destroy()

    win = tk.Tk()
    win.title("Cupom Fiscal")
    win.geometry("280x120")
    win.attributes("-topmost", True)
    tk.Label(win, text="Cliente deseja o cupom?", font=("Arial", 12)).pack(pady=10)
    btns = tk.Frame(win)
    btns.pack()
    tk.Button(btns, text="Sim", width=10, command=yes).grid(row=0, column=0, padx=10)
    tk.Button(btns, text="Não", width=10, command=no).grid(row=0, column=1, padx=10)
    win.mainloop()
    return result

def split_receipts(data):
    CUT_CMD = b'\x1D\x56'  # ESC/POS cut command
    parts = data.split(CUT_CMD)
    return [part + CUT_CMD for part in parts if part]

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

            wants = ask_receipt()

            if wants:
                print("Enviando tudo para a impressora.")
                print_to_real_printer(data)
            else:
                print("Cliente recusou cupom. Ignorando primeiras vias.")
                parts = split_receipts(data)
                if len(parts) > 2:
                    print_to_real_printer(parts[-1])  # Envia só a última via
                else:
                    print("Menos de 3 partes. Não imprimindo.")

if __name__ == "__main__":
    main()
