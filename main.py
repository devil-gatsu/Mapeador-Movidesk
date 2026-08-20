import tkinter as tk
from tkinter import messagebox
import pandas as pd
from playwright.sync_api import sync_playwright

# --- Interface Gráfica Compacta ---
def criar_interface():
    root = tk.Tk()
    root.title("Mapeador Movidesk")
    root.geometry("300x250") # Tamanho compacto
    root.attributes("-topmost", True) # Fica sempre por cima das outras janelas

    tk.Label(root, text="Cole o Cookie de Sessão:").pack(pady=5)
    cookie_entry = tk.Entry(root, width=40)
    cookie_entry.pack(pady=5)

    def iniciar_fase1():
        # Lógica da Fase 1 virá aqui
        status_label.config(text="Status: Mapeando Campos...")
        
    def iniciar_fase2():
        # Lógica da Fase 2 virá aqui
        status_label.config(text="Status: Mapeando Regras...")

    tk.Button(root, text="1. Iniciar Mapeamento de Campos", command=iniciar_fase1).pack(pady=5)
    tk.Button(root, text="2. Iniciar Mapeamento de Regras", command=iniciar_fase2).pack(pady=5)
    
    status_label = tk.Label(root, text="Status: Aguardando...", fg="blue")
    status_label.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()
