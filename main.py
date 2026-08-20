import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import queue
import time
import pandas as pd
from playwright.sync_api import sync_playwright

# Fila para comunicação entre a interface e o navegador
cmd_queue = queue.Queue()
global_df = None
global_filepath = ""

def playwright_worker(status_label):
    global global_df, global_filepath
    
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=False, args=["--start-maximized"])
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        
        # Vai direto para a tela de regras
        page.goto("https://omne.movidesk.com/CustomFieldRule")
        
        status_label.config(text="Status: Navegador aberto! Logue e abra uma Regra.", fg="green")
        
        # Fica escutando os cliques do seu botão "Capturar"
        while True:
            try:
                cmd = cmd_queue.get(timeout=1)
                
                if cmd == "CAPTURAR":
                    status_label.config(text="Status: Lendo a tela e salvando...", fg="orange")
                    
                    # 1. Pega o nome da regra (Tenta pegar do input de texto superior do Movidesk)
                    try:
                        nome_regra = page.evaluate("document.querySelector('input[type=\"text\"]').value")
                    except:
                        nome_regra = ""
                        
                    if not nome_regra or nome_regra.strip() == "":
                        nome_regra = "Regra_Desconhecida"
                        
                    # 2. Pega todo o texto visível na página atual
                    conteudo_tela = page.inner_text("body")
                    
                    # 3. Prepara as colunas da planilha (Índice 1 = Coluna B | Índice 3 = Coluna D)
                    coluna_nome = global_df.columns[1] 
                    coluna_alvo = global_df.columns[3] 
                    
                    global_df[coluna_alvo] = global_df[coluna_alvo].fillna("").astype(str)
                    campos_encontrados = 0
                    
                    # 4. Cruza a planilha com a tela
                    for idx, row in global_df.iterrows():
                        nome_campo = str(row[coluna_nome]).strip()
                        
                        # Se o nome do campo da planilha estiver visível na tela da regra
                        if nome_campo and nome_campo in conteudo_tela:
                            valor_atual = str(global_df.at[idx, coluna_alvo]).strip()
                            
                            # Escreve a regra nova
                            if not valor_atual or valor_atual.lower() == "nan":
                                global_df.at[idx, coluna_alvo] = nome_regra
                            else:
                                # Adiciona quebra de linha se o campo estiver em mais de uma regra
                                lista_regras = [r.strip() for r in valor_atual.split("\n")]
                                if nome_regra not in lista_regras:
                                    global_df.at[idx, coluna_alvo] = valor_atual + "\n" + nome_regra
                            
                            campos_encontrados += 1
                                    
                    # 5. Salva o Excel imediatamente
                    global_df.to_excel(global_filepath, index=False)
                    status_label.config(text=f"Status: '{nome_regra[:15]}...' salva! ({campos_encontrados} campos)", fg="green")
                    
            except queue.Empty:
                # Se o usuário fechar o Chrome manualmente, encerra a automação
                if page.is_closed():
                    status_label.config(text="Status: Navegador foi fechado.", fg="red")
                    break

def iniciar_navegador(status_label):
    global global_df, global_filepath
    
    filepath = filedialog.askopenfilename(
        title="Selecione a planilha de campos exportada",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if not filepath:
        return
        
    global_filepath = filepath
    try:
        global_df = pd.read_excel(filepath)
        # Garante que existam no mínimo as colunas A, B, C e D
        while len(global_df.columns) < 4:
            global_df[f"Coluna_{len(global_df.columns)}"] = ""
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler a planilha:\n{str(e)}")
        return
        
    status_label.config(text="Status: Iniciando Chrome...", fg="orange")
    
    # Inicia a thread do Playwright
    threading.Thread(target=playwright_worker, args=(status_label,), daemon=True).start()

def capturar_tela(status_label):
    # Envia o comando para a fila do Playwright trabalhar
    if global_df is not None:
        cmd_queue.put("CAPTURAR")
    else:
        messagebox.showwarning("Aviso", "Inicie o navegador e selecione a planilha primeiro!")

def criar_interface():
    root = tk.Tk()
    root.title("Assistente Movidesk")
    root.geometry("350x220")
    root.attributes("-topmost", True) # Fica sempre no topo para facilitar o clique
    root.resizable(False, False)

    tk.Label(root, text="Mapeador Semiautomático", font=("Arial", 11, "bold")).pack(pady=10)

    status_label = tk.Label(root, text="Status: Aguardando...", fg="blue", font=("Arial", 9))
    status_label.pack(pady=5)

    btn_iniciar = tk.Button(root, text="1. Iniciar Chrome e Abrir Planilha", command=lambda: iniciar_navegador(status_label), width=35, bg="#e1f5fe")
    btn_iniciar.pack(pady=5)

    # Botão de destaque para a ação repetitiva
    btn_capturar = tk.Button(root, text="2. CAPTURAR REGRA ATUAL", command=lambda: capturar_tela(status_label), width=35, height=2, bg="#c8e6c9", font=("Arial", 9, "bold"))
    btn_capturar.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()
