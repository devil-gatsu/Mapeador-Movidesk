import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import pandas as pd
import os
import time
from playwright.sync_api import sync_playwright

global_df = None
global_filepath = ""

def preparar_chrome(status_label):
    resposta = messagebox.askyesno(
        "Preparar Navegador", 
        "Isso vai forçar o encerramento de todos os processos invisíveis do Chrome e abri-lo corretamente para o robô.\n\nSalve o que estiver fazendo no navegador. Deseja continuar?"
    )
    if resposta:
        status_label.config(text="Status: Reiniciando Chrome...", fg="orange")
        try:
            # Mata o Chrome de verdade (processos fantasmas)
            os.system("taskkill /F /IM chrome.exe /T")
            time.sleep(1.5)
            # Abre o Chrome já com a porta certa
            os.system("start chrome.exe --remote-debugging-port=9222")
            status_label.config(text="Status: Chrome pronto! Pode logar no Movidesk.", fg="green")
        except Exception as e:
            messagebox.showerror("Erro", f"Não consegui abrir o Chrome: {e}")

def selecionar_planilha(status_label):
    global global_df, global_filepath
    
    filepath = filedialog.askopenfilename(
        title="Selecione a planilha de campos",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if not filepath:
        return
        
    try:
        global_filepath = filepath
        global_df = pd.read_excel(filepath)
        
        while len(global_df.columns) < 4:
            global_df[f"Coluna_Vazia_{len(global_df.columns)}"] = ""
            
        status_label.config(text=f"Status: Planilha carregada! ({len(global_df)} linhas)", fg="green")
    except Exception as e:
        status_label.config(text="Status: Erro ao ler planilha.", fg="red")
        messagebox.showerror("Erro", f"Erro ao ler a planilha:\n{str(e)}")

def executar_captura(status_label):
    global global_df, global_filepath
    
    if global_df is None:
        messagebox.showwarning("Aviso", "Selecione a planilha primeiro!")
        return
        
    status_label.config(text="Status: Conectando ao seu Chrome...", fg="orange")
    
    try:
        with sync_playwright() as p:
            # Tenta conectar pelo IP explícito para evitar bugs do localhost no Windows
            try:
                browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            except Exception as e1:
                try:
                    browser = p.chromium.connect_over_cdp("http://localhost:9222")
                except Exception as e2:
                    raise Exception(f"Erro de Conexão. O Chrome não abriu a porta 9222.\n\nPor favor, clique no Botão 1 para o sistema forçar a abertura correta.\n\nDetalhe do erro: {str(e1)}")
            
            context = browser.contexts[0]
            
            page = None
            for p_tab in context.pages:
                if "movidesk.com" in p_tab.url:
                    page = p_tab
                    break
                    
            if not page:
                raise Exception("A conexão deu certo, mas não encontrei a aba do Movidesk aberta!")

            status_label.config(text="Status: Lendo a tela...", fg="orange")

            nome_regra = page.evaluate("""() => {
                let inputs = Array.from(document.querySelectorAll('input[type="text"]'));
                let validInputs = inputs.filter(i => 
                    i.value.trim() !== '' && 
                    !i.placeholder.toLowerCase().includes('pesquisar') &&
                    !i.className.toLowerCase().includes('search')
                );
                if (validInputs.length > 0) return validInputs[0].value.trim();
                
                let title = document.querySelector('.k-window-title, h1, h2');
                if (title) return title.innerText.trim();
                
                return "Regra_Desconhecida";
            }""")
            
            if not nome_regra or nome_regra == "":
                nome_regra = "Regra_Desconhecida"

            textos_brutos = page.locator("body").inner_text().split('\n')
            textos_tela = [t.strip() for t in textos_brutos if t.strip() != ""]

            coluna_nome = global_df.columns[1] 
            coluna_alvo = global_df.columns[3] 
            global_df[coluna_alvo] = global_df[coluna_alvo].fillna("").astype(str)
            
            campos_encontrados = 0

            for idx, row in global_df.iterrows():
                nome_campo_planilha = str(row[coluna_nome]).strip()
                
                if not nome_campo_planilha or nome_campo_planilha.lower() == "nan":
                    continue
                
                match_encontrado = False
                for texto_na_tela in textos_tela:
                    if texto_na_tela == nome_campo_planilha or texto_na_tela.startswith(nome_campo_planilha):
                        match_encontrado = True
                        break
                        
                if match_encontrado:
                    valor_atual = str(global_df.at[idx, coluna_alvo]).strip()
                    
                    if not valor_atual or valor_atual.lower() == "nan":
                        global_df.at[idx, coluna_alvo] = nome_regra
                    else:
                        lista_regras = [r.strip() for r in valor_atual.split("\n")]
                        if nome_regra not in lista_regras:
                            global_df.at[idx, coluna_alvo] = valor_atual + "\n" + nome_regra
                            
                    campos_encontrados += 1
                            
            global_df.to_excel(global_filepath, index=False)
            status_label.config(text=f"Status: '{nome_regra[:15]}...' salva! ({campos_encontrados} campos)", fg="green")

    except Exception as e:
        status_label.config(text="Status: Erro ao capturar.", fg="red")
        messagebox.showerror("Erro", str(e))

def btn_capturar_thread(status_label):
    threading.Thread(target=executar_captura, args=(status_label,)).start()

def criar_interface():
    root = tk.Tk()
    root.title("Assistente Movidesk")
    root.geometry("350x230")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    tk.Label(root, text="Mapeador Semiautomático - CDP", font=("Arial", 11, "bold")).pack(pady=10)

    status_label = tk.Label(root, text="Status: Aguardando...", fg="blue", font=("Arial", 9))
    status_label.pack(pady=5)

    btn_preparar = tk.Button(root, text="1. Forçar Abertura do Chrome", command=lambda: preparar_chrome(status_label), width=35, bg="#fff9c4")
    btn_preparar.pack(pady=5)

    btn_planilha = tk.Button(root, text="2. Selecionar Planilha", command=lambda: selecionar_planilha(status_label), width=35, bg="#e1f5fe")
    btn_planilha.pack(pady=5)

    btn_capturar = tk.Button(root, text="3. CAPTURAR REGRA ATUAL", command=lambda: btn_capturar_thread(status_label), width=35, height=2, bg="#c8e6c9", font=("Arial", 9, "bold"))
    btn_capturar.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()
