import tkinter as tk
from tkinter import messagebox
import threading
import pandas as pd
from playwright.sync_api import sync_playwright

def executar_fase1(cookie_val, status_label):
    if not cookie_val:
        messagebox.showwarning("Aviso", "Por favor, insira o Cookie de Sessão do Movidesk!")
        return
    
    status_label.config(text="Status: Iniciando Fase 1 (Campos)...", fg="orange")
    
    try:
        # Aqui entrará a lógica do Playwright para:
        # 1. Abrir chromium invisível (headless=True)
        # 2. Injetar o cookie de sessão no domínio do Movidesk
        # 3. Ir na tela de campos adicionais, ler o total e fazer o scroll infinito
        # 4. Salvar o arquivo "mapeamento_movidesk.xlsx" com Id, Nome e Tipo
        
        # Simulação temporária para validação da interface:
        # (removeremos quando formos plugar a automação real dos campos)
        
        status_label.config(text="Status: Fase 1 Concluída! 1292 campos salvos.", fg="green")
        messagebox.sucesso = messagebox.showinfo("Sucesso", "Fase 1 finalizada com sucesso! Verifique a planilha.")
        
    except Exception as e:
        status_label.config(text="Status: Erro na execução.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

def executar_fase2(cookie_val, status_label):
    if not cookie_val:
        messagebox.showwarning("Aviso", "Por favor, insira o Cookie de Sessão do Movidesk!")
        return
        
    status_label.config(text="Status: Iniciando Fase 2 (Regras)...", fg="orange")
    
    try:
        # Aqui entrará a lógica do Playwright para:
        # 1. Ler o arquivo "mapeamento_movidesk.xlsx" gerado na Fase 1
        # 2. Acessar a tela de Regras de Exibição com o cookie injetado
        # 3. Clicar em cada regra, ir na aba Campos, cruzar os dados e salvar na coluna D
        
        status_label.config(text="Status: Fase 2 Concluída com sucesso!", fg="green")
        messagebox.showinfo("Sucesso", "Mapeamento de Regras finalizado!")
        
    except Exception as e:
        status_label.config(text="Status: Erro na execução.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

def criar_interface():
    root = tk.Tk()
    root.title("Mapeador Movidesk")
    root.geometry("320x240")
    root.attributes("-topmost", True) # Fica sempre visível/no topo
    root.resizable(False, False)

    tk.Label(root, text="Cole o Cookie de Sessão do Chrome:", font=("Arial", 9, "bold")).pack(pady=5)
    
    cookie_entry = tk.Entry(root, width=45)
    cookie_entry.pack(pady=2)

    status_label = tk.Label(root, text="Status: Aguardando ação...", fg="blue", font=("Arial", 9))
    status_label.pack(pady=10)

    # Funções de gatilho em Thread para não travar a janela enquanto o robô roda
    def botao_fase1_click():
        cookie = cookie_entry.get().strip()
        threading.Thread(target=executar_fase1, args=(cookie, status_label)).start()

    def botao_fase2_click():
        cookie = cookie_entry.get().strip()
        threading.Thread(target=executar_fase2, args=(cookie, status_label)).start()

    btn1 = tk.Button(root, text="1. Iniciar Mapeamento de Campos", command=botao_fase1_click, width=35, bg="#e1f5fe")
    btn1.pack(pady=5)

    btn2 = tk.Button(root, text="2. Iniciar Mapeamento de Regras", command=botao_fase2_click, width=35, bg="#e8f5e9")
    btn2.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()
