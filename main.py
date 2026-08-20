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
            # Mantemos visível para você acompanhar e poder intervir se necessário
            browser = p.chromium.launch(channel="chrome", headless=False, args=["--start-maximized"])
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
            
            # Injeção exata do cookie utilizando a URL para evitar erros de domínio (.omne vs omne)
            context.add_cookies([
                {"name": ".ASPXAUTH", "value": cookie_val.strip(), "url": BASE_URL}
            ])
            
            page = context.new_page()

            status_label.config(text="Status: [Fase 1] Acessando tela de Campos...", fg="orange")
            page.goto(f"{BASE_URL}/CustomField", timeout=60000)
            
            # BLINDAGEM: Se o cookie falhar ou pedir validação, ele pausa e deixa você logar manualmente na tela!
            if "login" in page.url.lower() or "account" in page.url.lower():
                status_label.config(text="Status: Faça o login na janela do Chrome...", fg="red")
                # O robô aguarda até 2 minutos para você digitar a senha e entrar na tela correta
                page.wait_for_url("**/CustomField**", timeout=120000)
                status_label.config(text="Status: Login detectado! Retomando automação...", fg="orange")

            status_label.config(text="Status: [Fase 1] Aguardando a tabela carregar...", fg="orange")
            
            # Espera a tabela ou o texto de totalização aparecer (com fallback resiliente)
            try:
                page.locator("text=total de").first.wait_for(state="visible", timeout=15000)
            except Exception:
                try:
                    page.locator("table tbody tr").first.wait_for(state="visible", timeout=15000)
                except Exception:
                    raise Exception("A página de Campos carregou, mas a tabela de dados está vazia ou bloqueada.")

            total_registros = 1292
            try:
                texto_paginacao = page.locator("text=total de").first.inner_text()
                match = re.search(r'total de (\d+)', texto_paginacao, re.IGNORECASE)
                if match:
                    total_registros = int(match.group(1))
            except:
                pass

            status_label.config(text=f"Status: [Fase 1] Rolando {total_registros} campos internamente...", fg="orange")

            ultimo_count = 0
            tentativas_paradas = 0

            for _ in range(100):
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
                    if tentativas_paradas >= 6: 
                        break
                else:
                    tentativas_paradas = 0
                    ultimo_count = linhas_atuais

            status_label.config(text="Status: [Fase 1] Extraindo colunas...", fg="orange")

            dados_campos = []
            linhas = page.locator("table tbody tr").all()

            for linha in linhas:
                colunas = linha.locator("td").all()
                
                if len(colunas) >= 4:
                    c_id = colunas[1].inner_text().strip()
                    c_nome = colunas[2].inner_text().strip()
                    c_tipo = colunas[3].inner_text().strip()

                    if c_id.isdigit():
                        dados_campos.append({
                            "Id": c_id,
                            "Nome": c_nome,
                            "Tipo": c_tipo,
                            "Regra de exibição": ""
                        })

            browser.close()

            if not dados_campos:
                raise Exception("Nenhum dado capturado. A tabela estava na tela, mas o robô não conseguiu ler as células.")

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
            browser = p.chromium.launch(channel="chrome", headless=False, args=["--start-maximized"])
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
            
            context.add_cookies([
                {"name": ".ASPXAUTH", "value": cookie_val.strip(), "url": BASE_URL}
            ])
            
            page = context.new_page()
            status_label.config(text="Status: [Fase 2] Acessando Regras de Exibição...", fg="orange")
            
            page.goto(f"{BASE_URL}/CustomFieldRule", timeout=60000)
            
            # Mesma blindagem para a Fase 2: se precisar logar, ele aguarda.
            if "login" in page.url.lower() or "account" in page.url.lower():
                status_label.config(text="Status: Faça o login na janela do Chrome...", fg="red")
                page.wait_for_url("**/CustomFieldRule**", timeout=120000)
                status_label.config(text="Status: Login detectado! Retomando automação...", fg="orange")
            
            try:
                page.locator("table tbody tr").first.wait_for(state="visible", timeout=20000)
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
                        page.locator("table tbody tr").first.wait_for(state="visible", timeout=20000)
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
