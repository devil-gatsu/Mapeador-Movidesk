import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# Variáveis Globais
df_atual = None
caminho_atual = ""
lista_campos = []
lista_regras = []

def carregar_dados_memoria(filepath):
    global df_atual, caminho_atual, lista_campos, lista_regras
    try:
        caminho_atual = filepath
        df_atual = pd.read_excel(filepath)
        
        # Garante colunas mínimas
        for col in ["Id", "Nome", "Tipo", "Regra de exibição"]:
            if col not in df_atual.columns:
                df_atual[col] = ""
                
        df_atual["Regra de exibição"] = df_atual["Regra de exibição"].fillna("").astype(str)
        df_atual["Nome"] = df_atual["Nome"].fillna("").astype(str)
        
        # Alimenta o Autocomplete
        lista_campos = sorted(list(df_atual["Nome"].unique()))
        todas_regras = df_atual["Regra de exibição"].str.split('\n').explode().str.strip().unique()
        lista_regras = sorted([r for r in todas_regras if r])
        
        lbl_status.config(text=f"Planilha carregada: {os.path.basename(filepath)} ({len(df_atual)} campos)", fg="green")
        atualizar_comboboxes()
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao carregar: {str(e)}")

# ================= FASE 1: CRUZAMENTO =================
def sincronizar_planilhas():
    nova_path = filedialog.askopenfilename(title="1. Selecione a Planilha NOVA (Regras Vazias)", filetypes=[("Excel", "*.xlsx")])
    if not nova_path: return
    
    antiga_path = filedialog.askopenfilename(title="2. Selecione a Planilha ANTIGA (Regras Preenchidas)", filetypes=[("Excel", "*.xlsx")])
    if not antiga_path: return
    
    try:
        df_nova = pd.read_excel(nova_path)
        df_antiga = pd.read_excel(antiga_path)
        
        # Cria um dicionário com Nome -> Regra da planilha antiga
        dict_regras = dict(zip(df_antiga['Nome'], df_antiga['Regra de exibição']))
        
        # Mapeia para a nova
        df_nova['Regra de exibição'] = df_nova['Nome'].map(dict_regras).fillna("")
        df_nova.to_excel(nova_path, index=False)
        
        messagebox.showinfo("Sucesso", "Planilhas cruzadas com sucesso! As regras foram copiadas.")
        carregar_dados_memoria(nova_path) # Já carrega para a Fase 2
    except Exception as e:
        messagebox.showerror("Erro", f"Falha no cruzamento: {str(e)}")

# ================= FASE 2: GERENCIADOR (CRUD) =================
def atualizar_comboboxes():
    cb_campo['values'] = lista_campos
    cb_regra['values'] = lista_regras

def autocomplete(event, cb, lista):
    digitado = cb.get().lower()
    if digitado == '':
        cb['values'] = lista
    else:
        filtrado = [item for item in lista if str(item).lower().startswith(digitado) or digitado in str(item).lower()]
        cb['values'] = filtrado

def mudar_interface():
    acao = var_acao.get()
    
    # Esconde tudo primeiro
    frame_alvo.pack_forget()
    frame_novo_valor.pack_forget()
    frame_criar.pack_forget()
    
    if acao in ["Excluir", "Alterar"]:
        frame_alvo.pack(fill="x", pady=5)
        if acao == "Alterar":
            frame_novo_valor.pack(fill="x", pady=5)
    elif acao == "Criar":
        frame_criar.pack(fill="x", pady=5)

def executar_acao():
    global df_atual
    if df_atual is None:
        messagebox.showwarning("Aviso", "Carregue ou sincronize uma planilha primeiro!")
        return
        
    acao = var_acao.get()
    alvo = var_alvo.get()
    campo_atual = cb_campo.get().strip()
    regra_atual = cb_regra.get().strip()
    
    try:
        if acao == "Excluir":
            if alvo == "Campo":
                df_atual = df_atual[df_atual["Nome"] != campo_atual]
            elif alvo == "Regra":
                # Limpa a regra específica onde ela existir
                df_atual["Regra de exibição"] = df_atual["Regra de exibição"].str.replace(regra_atual, "").str.strip()
                
        elif acao == "Alterar":
            novo_valor = entry_novo_valor.get().strip()
            if not novo_valor:
                messagebox.showwarning("Aviso", "Digite o novo valor!")
                return
                
            if alvo == "Campo":
                df_atual.loc[df_atual["Nome"] == campo_atual, "Nome"] = novo_valor
            elif alvo == "Regra":
                df_atual["Regra de exibição"] = df_atual["Regra de exibição"].str.replace(regra_atual, novo_valor)
                
        elif acao == "Criar":
            novo_id = entry_id.get().strip()
            novo_tipo = entry_tipo.get().strip()
            if not campo_atual:
                messagebox.showwarning("Aviso", "Digite o nome do Campo para criar!")
                return
                
            nova_linha = {"Id": novo_id, "Nome": campo_atual, "Tipo": novo_tipo, "Regra de exibição": regra_atual}
            df_atual = pd.concat([df_atual, pd.DataFrame([nova_linha])], ignore_index=True)

        # Salva e recarrega
        df_atual.to_excel(caminho_atual, index=False)
        carregar_dados_memoria(caminho_atual)
        messagebox.showinfo("Sucesso", f"Ação '{acao}' realizada e planilha atualizada!")
        
        # Limpa os campos
        cb_campo.set('')
        cb_regra.set('')
        entry_novo_valor.delete(0, tk.END)
        entry_id.delete(0, tk.END)
        entry_tipo.delete(0, tk.END)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao executar: {str(e)}")

# ================= INTERFACE GRÁFICA =================
root = tk.Tk()
root.title("Gerenciador de Regras Movidesk")
root.geometry("450x550")
root.resizable(False, False)

# --- Cabeçalho e Fase 1 ---
tk.Label(root, text="1. Sincronização Inicial", font=("Arial", 10, "bold")).pack(pady=(10, 0))
tk.Button(root, text="Cruzar Planilha Nova com Antiga", command=sincronizar_planilhas, bg="#e1f5fe").pack(fill="x", padx=20, pady=5)
tk.Button(root, text="Já tenho a planilha pronta (Carregar)", command=lambda: carregar_dados_memoria(filedialog.askopenfilename()), bg="#fff9c4").pack(fill="x", padx=20)

lbl_status = tk.Label(root, text="Nenhuma planilha carregada.", fg="red")
lbl_status.pack(pady=5)

tk.ttk.Separator(root, orient='horizontal').pack(fill='x', pady=10, padx=10)

# --- Fase 2: Gerenciador ---
tk.Label(root, text="2. Gerenciador (Criar / Alterar / Excluir)", font=("Arial", 10, "bold")).pack(pady=5)

# Ação
frame_acao = tk.Frame(root)
frame_acao.pack(pady=5)
var_acao = tk.StringVar(value="Alterar")
tk.Radiobutton(frame_acao, text="Alterar", variable=var_acao, value="Alterar", command=mudar_interface).pack(side="left")
tk.Radiobutton(frame_acao, text="Excluir", variable=var_acao, value="Excluir", command=mudar_interface).pack(side="left")
tk.Radiobutton(frame_acao, text="Criar", variable=var_acao, value="Criar", command=mudar_interface).pack(side="left")

# Inputs Principais (Com Autocomplete)
frame_inputs = tk.Frame(root)
frame_inputs.pack(fill="x", padx=20, pady=5)
tk.Label(frame_inputs, text="Campo:").pack(anchor="w")
cb_campo = ttk.Combobox(frame_inputs)
cb_campo.pack(fill="x")
cb_campo.bind('<KeyRelease>', lambda e: autocomplete(e, cb_campo, lista_campos))

tk.Label(frame_inputs, text="Regra:").pack(anchor="w", pady=(5,0))
cb_regra = ttk.Combobox(frame_inputs)
cb_regra.pack(fill="x")
cb_regra.bind('<KeyRelease>', lambda e: autocomplete(e, cb_regra, lista_regras))

# --- Paineis Dinâmicos (Escondidos/Mostrados via mudar_interface) ---
frame_alvo = tk.Frame(root)
tk.Label(frame_alvo, text="O que deseja modificar?").pack(side="left", padx=20)
var_alvo = tk.StringVar(value="Regra")
tk.Radiobutton(frame_alvo, text="Campo", variable=var_alvo, value="Campo").pack(side="left")
tk.Radiobutton(frame_alvo, text="Regra", variable=var_alvo, value="Regra").pack(side="left")

frame_novo_valor = tk.Frame(root)
tk.Label(frame_novo_valor, text="Novo Valor:").pack(anchor="w", padx=20)
entry_novo_valor = tk.Entry(frame_novo_valor)
entry_novo_valor.pack(fill="x", padx=20)

frame_criar = tk.Frame(root)
tk.Label(frame_criar, text="ID do Campo:").pack(anchor="w", padx=20)
entry_id = tk.Entry(frame_criar)
entry_id.pack(fill="x", padx=20)
tk.Label(frame_criar, text="Tipo do Campo:").pack(anchor="w", padx=20, pady=(5,0))
entry_tipo = tk.Entry(frame_criar)
entry_tipo.pack(fill="x", padx=20)

# Botão Executar
tk.Button(root, text="EXECUTAR AÇÃO", command=executar_acao, bg="#c8e6c9", font=("Arial", 9, "bold")).pack(fill="x", padx=20, pady=15)

mudar_interface() # Inicializa o layout dinâmico
root.mainloop()
