import tkinter as tk
from tkinter import messagebox
import threading
import time
import os
import re
import json
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
        dados_capturados_rede = []

        with sync_playwright() as p:
            browser = p.chromium.launch(channel="chrome", headless=True, args=["--start-maximized"])
            context = browser.new_context(viewport={"width": 1920, "height": 1080})
            
            # Injeta o cookie no subdomínio e no domínio raiz
            context.add_cookies([
                {"name": ".ASPXAUTH", "value": cookie_val.strip(), "domain": "omne.movidesk.com", "path": "/"},
                {"name": ".ASPXAUTH", "value": cookie_val.strip(), "domain": ".movidesk.com", "path": "/"}
            ])
            
            page = context.new_page()

            # Interceptador de respostas JSON da API do Movidesk
            def interceptar_resposta(response):
                try:
                    if "customfield" in response.url.lower() or "field" in response.url.lower():
                        content_type = response.headers.get("content-type", "")
                        if "application/json" in content_type:
                            data = response.json()
                            lista = []
                            if isinstance(data, list):
                                lista = data
                            elif isinstance(data, dict):
                                for k in ["data", "items", "rows", "CustomFields", "Result", "Data"]:
                                    if k in data and isinstance(data[k], list):
                                        lista = data[k]
                                        break
                            
                            for item in lista:
                                if isinstance(item, dict):
                                    c_id = str(item.get("id") or item.get("Id") or item.get("customFieldId") or "")
                                    c_nome = str(item.get("name") or item.get("Name") or item.get("description") or item.get("Description") or "")
                                    c_tipo = str(item.get("fieldTypeDescription") or item.get("typeDescription") or item.get("fieldType") or item.get("Type") or item.get("tipo") or "Personalizado")
                                    if c_nome:
                                        dados_capturados_rede.append({"Id": c_id, "Nome": c_nome, "Tipo": c_tipo, "Regra de exibição": ""})
                except Exception:
                    pass

            page.on("response", interceptar_resposta)

            status_label.config(text="Status: [Fase 1] Acessando tela de Campos...", fg="orange")
            page.goto(f"{BASE_URL}/Settings/CustomFields", timeout=60000)
            page.wait_for_load_state("networkidle")
            
            if "login" in page.url.lower():
                raise Exception("Sessão expirada ou Cookie inválido. O Movidesk redirecionou para o login.")

            time.sleep(4)
            status_label.config(text="Status: [Fase 1] Identificando total e rolando tabela...", fg="orange")

            # Busca total de registros no texto da página
            total_registros = 1292
            try:
                texto_pagina = page.locator("body").inner_text()
                match = re.search(r'total de\s+(\d+)', texto_pagina, re.IGNORECASE)
                if match:
                    total_registros = int(match.group(1))
            except Exception:
                pass

            # Rola a página e todos os contêineres internos de scroll
            ultimo_count = 0
            tentativas_paradas = 0

            for _ in range(60):
                page.evaluate("""() => {
                    window.scrollTo(0, document.body.scrollHeight);
                    document.querySelectorAll('div, section, main').forEach(el => {
                        if (el.scrollHeight > el.clientHeight) {
                            el.scrollTop = el.scrollHeight;
                        }
                    });
                }""")
                page.keyboard.press("End")
                page.keyboard.press("PageDown")
                time.sleep(1.2)

                linhas_atuais = page.locator("table tbody tr, .k-grid-content tr, tr").count()
                if linhas_atuais >= total_registros:
                    break
                
                if linhas_atuais == ultimo_count:
                    tentativas_paradas += 1
                    if tentativas_paradas >= 8:
                        break
                else:
                    tentativas_paradas = 0
                    ultimo_count = linhas_atuais

            status_label.config(text="Status: [Fase 1] Extraindo dados da tela...", fg="orange")

            # Varredura DOM com fallback resiliente
            dados_dom = []
            linhas = page.locator("tr").all()

            for idx_linha, linha in enumerate(linhas):
                tds = linha.locator("td").all()
                if len(tds) < 2:
                    continue

                textos = [td.inner_text().strip() for td in tds]
                textos_uteis = [t for t in textos if t and t not in ["-", "Sim", "Não", "Ativo", "Inativo", "Ativos", "Inativos"]]

                if not textos_uteis:
                    continue

                # 1. Tenta pegar o ID via Link de Edição (href)
                c_id = ""
                links = linha.locator("a").all()
                for lk in links:
                    href = lk.get_attribute("href") or ""
                    match_id = re.search(r'/(\d+)(?:\?|$|/)', href)
                    if match_id:
                        c_id = match_id.group(1)
                        break

                # 2. Tenta pegar o ID em atributos da linha
                if not c_id:
                    for attr in ["data-id", "id", "data-uid"]:
                        val = linha.get_attribute(attr) or ""
                        if val.isdigit():
                            c_id = val
                            break

                # 3. Tenta localizar nas células de texto
                if not c_id:
                    for t in textos_uteis:
                        if t.isdigit():
                            c_id = t
                            break

                # Extrai Nome e Tipo
                c_nome = ""
                c_tipo = "Personalizado"

                # Remove o ID da lista de textos úteis para isolar Nome e Tipo
                textos_sem_id = [t for t in textos_uteis if t != c_id]
                
                if len(textos_sem_id) >= 1:
                    c_nome = textos_sem_id[0]
                if len(textos_sem_id) >= 2:
                    c_tipo = textos_sem_id[1]

                if not c_id:
                    c_id = str(idx_linha)

                if c_nome and c_nome.lower() not in ["nome", "tipo", "ações", "status"]:
                    dados_dom.append({
                        "Id": c_id,
                        "Nome": c_nome,
                        "Tipo": c_tipo,
                        "Regra de exibição": ""
                    })

            browser.close()

            # Consolidação: Prefere dados da API se capturados, senão usa DOM
            if len(dados_capturados_rede) > len(dados_dom):
                dados_finais = dados_capturados_rede
            else:
                dados_finais = dados_dom

            if not dados_finais:
                raise Exception("A tabela carregou, mas os campos não foram identificados. Verifique se a página é a de Campos Adicionais.")

            df = pd.DataFrame(dados_finais).drop_duplicates(subset=["Nome"])
            df.to_excel(EXCEL_FILE, index=False)

            status_label.config(text=f"Status: [Fase 1] Concluído! {len(df)} campos salvos.", fg="green")
            messagebox.showinfo("Sucesso", f"Fase 1 finalizada!\n{len(df)} campos registrados em {EXCEL_FILE}.")

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
            
            page.goto(f"{BASE_URL}/Settings/TicketRule", timeout=60000)
            page.wait_for_load_state("networkidle")
            
            if "login" in page.url.lower():
                raise Exception("Sessão expirada ou Cookie inválido.")

            time.sleep(3)
            
            # Localiza todas as linhas de regras
            regras_linhas = page.locator("table tbody tr, tr").all()
            total_regras = len(regras_linhas)
            
            status_label.config(text=f"Status: [Fase 2] Analisando {total_regras} regras...", fg="orange")
            
            for index in range(total_regras):
                try:
                    linhas_atuais = page.locator("table tbody tr, tr").all()
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
                    
                    # Clica na aba 'Campos'
                    aba_campos = page.locator("text=Campos, a:has-text('Campos')").first
                    if aba_campos.count() > 0:
                        aba_campos.click()
                        time.sleep(1)
                        
                        conteudo_tela = page.inner_text("body")
                        
                        # Cruza com os campos da planilha
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

                    # Retorna à listagem principal
                    botao_cancelar = page.locator("text=CANCELAR, button:has-text('Cancelar'), a:has-text('Cancelar')").first
                    if botao_cancelar.count() > 0:
                        botao_cancelar.click()
                    else:
                        page.goto(f"{BASE_URL}/Settings/TicketRule", timeout=30000)
                        
                    page.wait_for_load_state("networkidle")
                    time.sleep(1.2)
                    
                except Exception as ex_item:
                    try:
                        page.goto(f"{BASE_URL}/Settings/TicketRule", timeout=30000)
                        page.wait_for_load_state("networkidle")
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
