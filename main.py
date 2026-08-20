import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import time
import os
import pandas as pd
from playwright.sync_api import sync_playwright

BASE_URL = "https://omne.movidesk.com"
SESSION_DIR = os.path.join(os.getcwd(), "sessao_movidesk") # Pasta onde o robô vai salvar seu login

def abrir_navegador_para_login(status_label):
    status_label.config(text="Status: Abrindo navegador para você logar...", fg="orange")
    try:
        with sync_playwright() as p:
            # Cria um navegador que LEMBRA do login
            context = p.chromium.launch_persistent_context(
                user_data_dir=SESSION_DIR,
                channel="chrome",
                headless=False,
                args=["--start-maximized"]
            )
            page = context.pages[0]
            page.goto(f"{BASE_URL}/Account/Login")
            
            messagebox.showinfo("Login", "Faça o login normalmente na janela do Chrome que abriu.\nQuando terminar e estiver na tela inicial do Movidesk, pode fechar a janela do navegador!")
            
            # Mantém o navegador aberto até você fechá-lo
            page.wait_for_event("close", timeout=0)
            
        status_label.config(text="Status: Sessão salva! Pronto para mapear.", fg="green")
    except Exception as e:
        status_label.config(text="Status: Erro ao abrir navegador.", fg="red")
        messagebox.showerror("Erro", str(e))

def executar_mapeamento(filepath, status_label):
    status_label.config(text="Status: Iniciando automação das Regras...", fg="orange")
    
    try:
        # Lê a planilha que você exportou
        df = pd.read_excel(filepath)
        
        # Garante que a planilha tem pelo menos 4 colunas (A, B, C, D) para não dar erro
        while len(df.columns) < 4:
            df[f"Coluna_Vazia_{len(df.columns)}"] = ""
            
        # Define os alvos com base nas colunas físicas (0=A, 1=B, 2=C, 3=D)
        coluna_nome_campo = df.columns[1] # Coluna B
        coluna_regra_alvo = df.columns[3] # Coluna D
        
        # Garante que a Coluna D é tratada como texto
        df[coluna_regra_alvo] = df[coluna_regra_alvo].fillna("").astype(str)

        with sync_playwright() as p:
            # Abre usando a sessão salva (já logado)
            context = p.chromium.launch_persistent_context(
                user_data_dir=SESSION_DIR,
                channel="chrome",
                headless=False,
                args=["--start-maximized"]
            )
            
            page = context.pages[0]
            status_label.config(text="Status: Acessando Regras de Exibição...", fg="orange")
            
            page.goto(f"{BASE_URL}/CustomFieldRule", timeout=60000)
            
            if "login" in page.url.lower() or "account" in page.url.lower():
                raise Exception("O robô não está logado. Por favor, use o Botão 1 primeiro para fazer login.")
            
            try:
                page.locator("table tbody tr").first.wait_for(state="visible", timeout=20000)
            except:
                raise Exception("A listagem de regras não carregou.")
            
            regras_linhas = page.locator("table tbody tr").all()
            total_regras = len(regras_linhas)
            
            status_label.config(text=f"Status: Analisando {total_regras} regras...", fg="orange")
            
            for index in range(total_regras):
                try:
                    linhas_atuais = page.locator("table tbody tr").all()
                    if index >= len(linhas_atuais):
                        break
                        
                    linha_regra = linhas_atuais[index]
                    link_regra = linha_regra.locator("a").first
                    
                    if link_regra.count() == 0:
                        continue
                        
                    nome_regra = link_regra.inner_text().strip()
                    if not nome_regra:
                        nome_regra = f"Regra #{index+1}"

                    status_label.config(text=f"Status: Lendo regra {index+1}/{total_regras}...", fg="orange")
                    
                    link_regra.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(1.5)
                    
                    # Vai para a aba Campos
                    aba_campos = page.locator("text=Campos, a:has-text('Campos')").first
                    if aba_campos.count() > 0:
                        aba_campos.click()
                        time.sleep(1)
                        
                        # Lê tudo que está escrito na tela
                        conteudo_tela = page.inner_text("body")
                        
                        # Varre a sua planilha (Coluna B) e cruza com a tela
                        for idx_df, row in df.iterrows():
                            nome_campo = str(row[coluna_nome_campo]).strip()
                            
                            if nome_campo and nome_campo in conteudo_tela:
                                valor_atual = str(df.at[idx_df, coluna_regra_alvo]).strip()
                                
                                # Anota a regra na Coluna D
                                if not valor_atual or valor_atual.lower() == "nan":
                                    df.at[idx_df, coluna_regra_alvo] = nome_regra
                                else:
                                    lista_regras = [r.strip() for r in valor_atual.split("\n")]
                                    if nome_regra not in lista_regras:
                                        df.at[idx_df, coluna_regra_alvo] = valor_atual + "\n" + nome_regra

                    # Cancela para voltar à lista
                    botao_cancelar = page.locator("text=CANCELAR, button:has-text('Cancelar'), a:has-text('Cancelar')").first
                    if botao_cancelar.count() > 0:
                        botao_cancelar.click()
                    else:
                        page.goto(f"{BASE_URL}/CustomFieldRule", timeout=30000)
                        
                    page.wait_for_load_state("networkidle")
                    time.sleep(1.2)
                    
                except Exception as ex_item:
                    try:
                        page.goto(f"{BASE_URL}/CustomFieldRule", timeout=30000)
                        page.locator("table tbody tr").first.wait_for(state="visible", timeout=20000)
                    except Exception:
                        pass
                    continue

            # Sobrescreve o mesmo arquivo Excel que você abriu
            df.to_excel(filepath, index=False)
            context.close()
            
            status_label.config(text="Status: Mapeamento concluído com sucesso!", fg="green")
            messagebox.showinfo("Sucesso", f"Automação finalizada!\nO arquivo {os.path.basename(filepath)} foi atualizado com as regras na Coluna D.")
            
    except Exception as e:
        status_label.config(text="Status: Erro na automação.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

def criar_interface():
    root = tk.Tk()
    root.title("Mapeador Movidesk - Regras")
    root.geometry("350x220")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    tk.Label(root, text="Automação de Regras de Exibição", font=("Arial", 10, "bold")).pack(pady=10)

    status_label = tk.Label(root, text="Status: Aguardando ação...", fg="blue", font=("Arial", 9))
    status_label.pack(pady=5)

    def btn_login_click():
        threading.Thread(target=abrir_navegador_para_login, args=(status_label,)).start()

    def btn_mapear_click():
        # Abre janela para você selecionar o arquivo de relatório exportado
        filepath = filedialog.askopenfilename(
            title="Selecione a planilha de campos exportada",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if filepath:
            threading.Thread(target=executar_mapeamento, args=(filepath, status_label)).start()

    btn1 = tk.Button(root, text="1. Fazer Login no Movidesk", command=btn_login_click, width=35, bg="#fff9c4")
    btn1.pack(pady=5)

    btn2 = tk.Button(root, text="2. Selecionar Planilha e Mapear Regras", command=btn_mapear_click, width=35, bg="#e8f5e9")
    btn2.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()
