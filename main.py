import tkinter as tk
from tkinter import messagebox
import threading
import time
import os
import re
import pandas as pd
from playwright.sync_api import sync_playwright

EXCEL_FILE = "mapeamento_movidesk.xlsx"
BASE_URL = "https://omne.movidesk.com"

def executar_fase1(cookie_val, status_label):
    if not cookie_val:
        messagebox.showwarning("Aviso", "Por favor, insira o Cookie .ASPXAUTH do Movidesk!")
        return
    
    status_label.config(text="Status: [Fase 1] Conectando ao Movidesk...", fg="orange")
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(channel="chrome", headless=True, args=["--start-maximized"])
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
            
            context.add_cookies([
                {"name": ".ASPXAUTH", "value": cookie_val.strip(), "domain": "omne.movidesk.com", "path": "/"},
                {"name": ".ASPXAUTH", "value": cookie_val.strip(), "domain": ".movidesk.com", "path": "/"}
            ])
            
            page = context.new_page()

            status_label.config(text="Status: [Fase 1] Acessando tela de Campos...", fg="orange")
            page.goto(f"{BASE_URL}/CustomField", timeout=60000)
            
            # 1. Espera cirúrgica baseada na sua imagem ("Exibindo de 1 até...")
            status_label.config(text="Status: [Fase 1] Aguardando a tabela aparecer...", fg="orange")
            try:
                # O robô agora espera ESSE texto aparecer para ter certeza que carregou
                page.locator("text=total de").wait_for(state="visible", timeout=30000)
            except Exception:
                raise Exception("A página carregou, mas a lista de campos não apareceu. Verifique o cookie.")

            # 2. Descobre o total de registros lendo o texto do cabeçalho
            total_registros = 1292
            try:
                texto_paginacao = page.locator("text=total de").first.inner_text()
                match = re.search(r'total de (\d+)', texto_paginacao, re.IGNORECASE)
                if match:
                    total_registros = int(match.group(1))
            except:
                pass

            status_label.config(text=f"Status: [Fase 1] Rolando {total_registros} campos internamente...", fg="orange")

            # 3. Lógica de scroll INTERNO da tabela (foca no contêiner com a barra)
            ultimo_count = 0
            tentativas_paradas = 0

            for _ in range(100):
                # O comando JS agora mira explicitamente em áreas roláveis internas, como o k-grid-content do Movidesk
                page.evaluate("""() => {
                    let grid = document.querySelector('.k-grid-content') || 
                               document.querySelector('.table-responsive') || 
                               document.querySelector('table').parentElement;
                    if(grid) {
                        grid.scrollTop = grid.scrollHeight;
                    }
                }""")
                time.sleep(1.5)
                
                linhas_atuais = page.locator("table tbody tr").count()
                
                if linhas_atuais >= total_registros:
                    break
                    
                if linhas_atuais == ultimo_count:
                    tentativas_paradas += 1
                    if tentativas_paradas >= 6: # Se a tabela parar de crescer, encerra o loop
                        break
                else:
                    tentativas_paradas = 0
                    ultimo_count = linhas_atuais

            status_label.config(text="Status: [Fase 1] Extraindo colunas...", fg="orange")

            # 4. Extração exata baseada na estrutura da sua imagem
            dados_campos = []
            linhas = page.locator("table tbody tr").all()

            for linha in linhas:
                colunas = linha.locator("td").all()
                
                # A imagem mostra que existem várias colunas. Precisamos de pelo menos 4 (Check, ID, Nome, Tipo)
                if len(colunas) >= 4:
                    c_id = colunas[1].inner_text().strip()
                    c_nome = colunas[2].inner_text().strip()
                    c_tipo = colunas[3].inner_text().strip()

                    # Só salva se a coluna de ID realmente for um número (evita cabeçalhos perdidos)
                    if c_id.isdigit():
                        dados_campos.append({
                            "Id": c_id,
                            "Nome": c_nome,
                            "Tipo": c_tipo,
                            "Regra de exibição": ""
                        })

            browser.close()

            if not dados_campos:
                raise Exception("Nenhum dado capturado. A tabela foi encontrada, mas não foi possível ler as colunas de ID e Nome.")

            df = pd.DataFrame(dados_campos).drop_duplicates(subset=["Id"])
            df.to_excel(EXCEL_FILE, index=False)

            status_label.config(text=f"Status: [Fase 1] Concluído! {len(df)} campos salvos.", fg="green")
            messagebox.showinfo("Sucesso", f"Fase 1 finalizada com sucesso!\n{len(df)} registros salvos em {EXCEL_FILE}.")

    except Exception as e:
        status_label.config(text="Status: Erro na Fase 1.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro na Fase 1:\n{str(e)}")


def executar_fase2(cookie_val, status_label):
    if not cookie_val:
        messagebox.showwarning("Aviso", "Por favor, insira o Cookie .ASPXAUTH do Movidesk!")
        return
        
    if not os.path.exists(EXCEL_FILE):
        messagebox.showwarning("Aviso", "A planilha da Fase 1 não foi encontrada! Execute a Fase 1 primeiro.")
        return
        
    status_label.config(text="Status: [Fase 2] Carregando planilha e conectando...", fg="orange")
    
    try:
        df = pd.read_excel(EXCEL_FILE)
        if "Regra de exibição" not in df.columns:
            df["Regra de exibição"] = ""
        df["Regra de exibição"] = df["Regra de exibição"].fillna("").astype(str)

        with sync_playwright() as p:
            browser = p.chromium.launch(channel="chrome", headless=True, args=["--start-maximized"])
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
            
            context.add_cookies([
                {"name": ".ASPXAUTH", "value": cookie_val.strip(), "domain": "omne.movidesk.com", "path": "/"},
                {"name": ".ASPXAUTH", "value": cookie_val.strip(), "domain": ".movidesk.com", "path": "/"}
            ])
            
            page = context.new_page()
            status_label.config(text="Status: [Fase 2] Acessando Regras de Exibição...", fg="orange")
            
            page.goto(f"{BASE_URL}/CustomFieldRule", timeout=60000)
            
            try:
                page.locator("table tbody tr").first.wait_for(state="visible", timeout=30000)
            except:
                raise Exception("A listagem de regras não carregou corretamente.")
            
            regras_linhas = page.locator("table tbody tr").all()
            total_regras = len(regras_linhas)
            
            status_label.config(text=f"Status: [Fase 2] Analisando {total_regras} regras...", fg="orange")
            
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

                    status_label.config(text=f"Status: [Fase 2] Lendo regra {index+1}/{total_regras}...", fg="orange")
                    
                    link_regra.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(1.5)
                    
                    aba_campos = page.locator("text=Campos, a:has-text('Campos')").first
                    if aba_campos.count() > 0:
                        aba_campos.click()
                        time.sleep(1)
                        
                        conteudo_tela = page.inner_text("body")
                        
                        for idx_df, row in df.iterrows():
                            nome_campo = str(row["Nome"]).strip()
                            if nome_campo and nome_campo in conteudo_tela:
                                valor_atual = str(df.at[idx_df, "Regra de exibição"]).strip()
                                
                                if not valor_atual or valor_atual.lower() == "nan":
                                    df.at[idx_df, "Regra de exibição"] = nome_regra
                                else:
                                    lista_regras = [r.strip() for r in valor_atual.split("\n")]
                                    if nome_regra not in lista_regras:
                                        df.at[idx_df, "Regra de exibição"] = valor_atual + "\n" + nome_regra

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
                        page.locator("table tbody tr").first.wait_for(state="visible", timeout=30000)
                    except Exception:
                        pass
                    continue

            df.to_excel(EXCEL_FILE, index=False)
            browser.close()
            
            status_label.config(text="Status: [Fase 2] Concluído com sucesso!", fg="green")
            messagebox.showinfo("Sucesso", f"Fase 2 finalizada!\nPlanilha {EXCEL_FILE} atualizada com as regras.")
            
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
