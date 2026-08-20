import tkinter as tk
from tkinter import messagebox
import threading
import time
import os
import pandas as pd
from playwright.sync_api import sync_playwright

EXCEL_FILE = "mapeamento_movidesk.xlsx"

def obter_url_base(cookie_val):
    return "https://omne.movidesk.com"

def executar_fase1(cookie_val, status_label):
    if not cookie_val:
        messagebox.showwarning("Aviso", "Por favor, insira o Cookie .ASPXAUTH do Movidesk!")
        return
    
    status_label.config(text="Status: [Fase 1] Conectando ao Movidesk...", fg="orange")
    
    try:
        base_url = obter_url_base(cookie_val)
        
        with sync_playwright() as p:
            # Utiliza o Google Chrome instalado no sistema local, evitando erros de driver
            browser = p.chromium.launch(channel="chrome", headless=True, args=["--start-maximized"])
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
            
            # Injeta o cookie de sessão exato fornecido pelo usuário
            context.add_cookies([{
                "name": ".ASPXAUTH",
                "value": cookie_val.strip(),
                "domain": "omne.movidesk.com",
                "path": "/"
            }])
            
            page = context.new_page()
            status_label.config(text="Status: [Fase 1] Acessando tela de Campos Adicionais...", fg="orange")
            
            # Navega direto para a tela de Campos Adicionais do Movidesk
            page.goto(f"{base_url}/Settings/CustomFields", timeout=60000)
            page.wait_for_load_state("networkidle")
            
            if "login" in page.url.lower():
                raise Exception("Sessão expirada ou Cookie inválido. O Movidesk redirecionou para o login.")

            status_label.config(text="Status: [Fase 1] Lendo total de campos...", fg="orange")
            time.sleep(3)
            
            total_registros = 0
            try:
                texto_info = page.locator("text=total de").locator("xpath=..").inner_text()
                import re
                numeros = re.findall(r'\d+', texto_info)
                if len(numeros) >= 2:
                    total_registros = int(numeros[-1])
            except:
                total_registros = 1292

            status_label.config(text=f"Status: [Fase 1] Executando scroll para carregar {total_registros} campos...", fg="orange")
            
            ultimo_tamanho = 0
            tentativas_sem_mudanca = 0
            
            while True:
                page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1.5)
                
                linhas = page.locator("table tbody tr")
                tamanho_atual = linhas.count()
                
                if tamanho_atual >= total_registros or tentativas_sem_mudanca > 10:
                    break
                    
                if tamanho_atual == ultimo_tamanho:
                    tentativas_sem_mudanca += 1
                else:
                    tentativas_sem_mudanca = 0
                    ultimo_tamanho = tamanho_atual

            status_label.config(text="Status: [Fase 1] Extraindo dados dos campos...", fg="orange")
            
            dados_campos = []
            linhas = page.locator("table tbody tr")
            count = linhas.count()
            
            for i in range(count):
                linha = linhas.nth(i)
                colunas = linha.locator("td")
                if colunas.count() >= 3:
                    id_campo = colunas.nth(0).inner_text().strip()
                    nome_campo = colunas.nth(1).inner_text().strip()
                    tipo_campo = colunas.nth(2).inner_text().strip()
                    
                    if id_campo.isdigit():
                        dados_campos.append({
                            "Id": id_campo,
                            "Nome": nome_campo,
                            "Tipo": tipo_campo,
                            "Regra de exibição": ""
                        })

            browser.close()
            
            df = pd.DataFrame(dados_campos)
            df.to_excel(EXCEL_FILE, index=False)
            
            status_label.config(text=f"Status: [Fase 1] Concluído! {len(dados_campos)} campos salvos.", fg="green")
            messagebox.showinfo("Sucesso", f"Fase 1 finalizada! {len(dados_campos)} registros salvos em {EXCEL_FILE}.")
            
    except Exception as e:
        status_label.config(text="Status: Erro na Fase 1.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro na Fase 1:\n{str(e)}")

def executar_fase2(cookie_val, status_label):
    if not cookie_val:
        messagebox.showwarning("Aviso", "Por favor, insira o Cookie .ASPXAUTH do Movidesk!")
        return
        
    if not os.path.exists(EXCEL_FILE):
        messagebox.showwarning("Aviso", "O arquivo de Excel dos campos não foi encontrado! Execute a Fase 1 primeiro.")
        return
        
    status_label.config(text="Status: [Fase 2] Carregando planilha e conectando...", fg="orange")
    
    try:
        base_url = obter_url_base(cookie_val)
        df = pd.read_excel(EXCEL_FILE)
        if "Regra de exibição" not in df.columns:
            df["Regra de exibição"] = ""
        df["Regra de exibição"] = df["Regra de exibição"].fillna("").astype(str)

        with sync_playwright() as p:
            browser = p.chromium.launch(channel="chrome", headless=True, args=["--start-maximized"])
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
            
            context.add_cookies([{
                "name": ".ASPXAUTH",
                "value": cookie_val.strip(),
                "domain": "omne.movidesk.com",
                "path": "/"
            }])
            
            page = context.new_page()
            status_label.config(text="Status: [Fase 2] Acessando Regras de Exibição...", fg="orange")
            
            page.goto(f"{base_url}/Settings/TicketRule", timeout=60000)
            page.wait_for_load_state("networkidle")
            
            if "login" in page.url.lower():
                raise Exception("Sessão expirada ou Cookie inválido.")

            time.sleep(3)
            
            total_regras_elementos = page.locator("table tbody tr")
            total_regras = total_regras_elementos.count()
            
            status_label.config(text=f"Status: [Fase 2] Analisando {total_regras} regras...", fg="orange")
            
            for index in range(total_regras):
                try:
                    linhas_regras = page.locator("table tbody tr")
                    if index >= linhas_regras.count():
                        break
                        
                    linha_regra = linhas_regras.nth(index)
                    link_regra = linha_regra.locator("td a").first
                    if link_regra.count() == 0:
                        link_regra = linha_regra.locator("td").first
                        
                    nome_regra = link_regra.inner_text().strip()
                    if not nome_regra:
                        nome_regra = f"Regra #{index+1}"

                    status_label.config(text=f"Status: [Fase 2] Lendo regra {index+1}/{total_regras}...", fg="orange")
                    
                    link_regra.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(1.5)
                    
                    aba_campos = page.locator("text=Campos").first
                    aba_campos.click()
                    time.sleep(1)
                    
                    campos_encontrados_nesta_regra = []
                    pagina_texto_completo = page.inner_text("body")
                    
                    for idx, row in df.iterrows():
                        nome_f = str(row["Nome"]).strip()
                        if nome_f and nome_f in pagina_texto_completo:
                            campos_encontrados_nesta_regra.append(nome_f)

                    for nome_campo_encontrado in campos_encontrados_nesta_regra:
                        mask = df["Nome"] == nome_campo_encontrado
                        if mask.any():
                            valor_atual = str(df.loc[mask, "Regra de exibição"].values[0]).strip()
                            
                            if not valor_atual or valor_atual == "nan":
                                novo_valor = nome_regra
                            else:
                                if nome_regra not in valor_atual.split("\n"):
                                    novo_valor = valor_atual + "\n" + nome_regra
                                else:
                                    novo_valor = valor_atual
                                    
                            df.loc[mask, "Regra de exibição"] = novo_valor

                    botao_cancelar = page.locator("text=CANCELAR").first
                    if botao_cancelar.count() > 0:
                        botao_cancelar.click()
                    else:
                        page.go_back()
                        
                    page.wait_for_load_state("networkidle")
                    time.sleep(1.5)
                    
                except Exception as ex_regra:
                    print(f"Erro ao processar regra no índice {index}: {str(ex_regra)}")
                    try:
                        page.goto(f"{base_url}/Settings/TicketRule", timeout=30000)
                        page.wait_for_load_state("networkidle")
                    except:
                        pass
                    continue

            df.to_excel(EXCEL_FILE, index=False)
            browser.close()
            
            status_label.config(text="Status: [Fase 2] Mapeamento de Regras concluído!", fg="green")
            messagebox.showinfo("Sucesso", "Mapeamento de Regras finalizado com sucesso! Planilha atualizada.")
            
    except Exception as e:
        status_label.config(text="Status: Erro na Fase 2.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro na Fase 2:\n{str(e)}")

def criar_interface():
    root = tk.Tk()
    root.title("Mapeador Movidesk")
    root.geometry("320x240")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    tk.Label(root, text="Cole o Cookie .ASPXAUTH do Chrome:", font=("Arial", 9, "bold")).pack(pady=5)
    
    cookie_entry = tk.Entry(root, width=45)
    cookie_entry.pack(pady=2)

    status_label = tk.Label(root, text="Status: Aguardando ação...", fg="blue", font=("Arial", 9))
    status_label.pack(pady=10)

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
