import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.worksheet.protection import SheetProtection
import threading

# === Funções ===
def selecionar_pasta():
    caminho = filedialog.askdirectory()
    if caminho:
        pasta_entry.delete(0, tk.END)
        pasta_entry.insert(0, caminho)

def proteger_planilhas():
    pasta = pasta_entry.get()
    senha = senha_entry.get()

    if not os.path.isdir(pasta):
        messagebox.showerror("Erro", "Caminho de pasta inválido.")
        return

    if not senha:
        messagebox.showwarning("Aviso", "Digite uma senha para proteção.")
        return

    arquivos = [
        f for f in os.listdir(pasta)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    ]

    if not arquivos:
        messagebox.showinfo("Info", "Nenhum arquivo .xlsx encontrado.")
        return

    erros = []
    progress_bar["maximum"] = len(arquivos)
    progress_bar["value"] = 0

    def update_progress(i):
        progress_bar["value"] = i
        porcentagem = (i / len(arquivos)) * 100
        porcentagem_label.config(text=f"{int(porcentagem)}%")
        root.update_idletasks()

    # Função para simular o preenchimento gradual da barra
    def simulate_progress(i):
        if i <= len(arquivos):
            update_progress(i)
            root.after(50, simulate_progress, i + 1)
        else:
            if erros:
                messagebox.showwarning("Concluído com Erros", "\n".join(erros))
            else:
                if messagebox.askokcancel("Concluído", "Todas as planilhas foram protegidas com sucesso!"):
                    reset_interface()

    for i, arquivo in enumerate(arquivos):
        caminho_arquivo = os.path.join(pasta, arquivo)
        try:
            wb = load_workbook(caminho_arquivo)
            for ws in wb.worksheets:
                ws.protection = SheetProtection(sheet=True, password=senha)
            wb.save(caminho_arquivo)
        except PermissionError:
            erros.append(f"{arquivo}: Arquivo está aberto. Feche-o para aplicar proteção.")
        except Exception as e:
            erros.append(f"{arquivo}: Erro - {str(e)}")

    # Iniciar o preenchimento gradual da barra
    root.after(500, simulate_progress, 1)  # Atraso inicial de 500ms para começar o preenchimento gradual

def iniciar_protecao_thread():
    threading.Thread(target=proteger_planilhas).start()

def reset_interface():
    # Limpar os campos
    pasta_entry.delete(0, tk.END)
    senha_entry.delete(0, tk.END)
    
    # Zerando a barra de progresso
    progress_bar["value"] = 0
    porcentagem_label.config(text="0%")

# === Interface ===
root = tk.Tk()
root.title("Proteger Planilhas Excel")
root.geometry("500x300")
root.configure(bg="#f0f2f5")

style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=("Segoe UI", 10), padding=6)
style.configure("TLabel", font=("Segoe UI", 10), background="#f0f2f5")
style.configure("TEntry", padding=5)

frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

# Campo da pasta
ttk.Label(frame, text="Caminho da pasta com planilhas:").grid(row=0, column=0, sticky="w")
pasta_entry = ttk.Entry(frame, width=50)
pasta_entry.grid(row=1, column=0, padx=(0, 10), pady=5, sticky="w")
ttk.Button(frame, text="Selecionar", command=selecionar_pasta).grid(row=1, column=1, pady=5)

# Campo da senha
ttk.Label(frame, text="Senha de proteção:").grid(row=2, column=0, sticky="w", pady=(10, 0))
senha_entry = ttk.Entry(frame, show="*", width=30)
senha_entry.grid(row=3, column=0, pady=5, sticky="w")

# Botão principal
ttk.Button(frame, text="Proteger Planilhas", command=iniciar_protecao_thread).grid(row=4, column=0, columnspan=2, pady=20)

# Barra de progresso
progress_bar = ttk.Progressbar(frame, length=400, mode="determinate", style="green.Horizontal.TProgressbar")
progress_bar.grid(row=5, column=0, columnspan=2, pady=(0, 10))

# Adicionar a porcentagem
porcentagem_label = ttk.Label(frame, text="0%", font=("Segoe UI", 10))
porcentagem_label.grid(row=6, column=0, columnspan=2)

# Criar o estilo para a barra de progresso verde
style.configure("green.Horizontal.TProgressbar", thickness=20, background="#4CAF50")

# Inicia a interface
root.mainloop()
