import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ==========================================
# FUNÇÕES DE LÓGICA (BACK-END)
# ==========================================

def executar_processamento():
    caminho_entrada = var_entrada.get()
    pasta_destino = var_saida.get()
    nome_base = var_nome_arquivo.get()

    if not caminho_entrada or not os.path.exists(caminho_entrada):
        messagebox.showerror("Erro", "Selecione um arquivo de entrada válido.")
        return
    if not pasta_destino or not os.path.exists(pasta_destino):
        messagebox.showerror("Erro", "Selecione uma pasta de saída válida.")
        return

    # 1. Capturar as regras da interface
    mapa_horas = {}
    for bloco in lista_tipos:
        # Se for o bloco Padrão, forçamos o tipo a ser vazio ("") para funcionar como coringa
        tipo = "" if bloco.get('is_padrao') else bloco['entry_tipo'].get().strip().lower()
        
        for linha in bloco['linhas']:
            comp = linha['comp'].get().strip().lower()
            hora_str = linha['hora'].get().strip()

            # Validação Obrigatória do Grupo Padrão
            if bloco.get('is_padrao') and not hora_str:
                messagebox.showerror("Atenção", f"O preenchimento de horas para a complexidade '{comp}' no Grupo Padrão é obrigatório.")
                return

            if comp and hora_str:
                try:
                    mapa_horas[(tipo, comp)] = float(hora_str)
                except ValueError:
                    nome_exibicao = tipo if tipo else "Grupo Padrão"
                    messagebox.showerror("Erro", f"Hora inválida ('{hora_str}') na complexidade '{comp}' do '{nome_exibicao}'.")
                    return

    if not mapa_horas:
        messagebox.showerror("Erro", "Configure pelo menos uma regra válida de horas.")
        return

    # 2. Versionamento automático
    versao = 1
    while True:
        nome_arquivo = f"{nome_base}{versao}.xlsx"
        caminho_final = os.path.join(pasta_destino, nome_arquivo)
        if not os.path.exists(caminho_final):
            break
        versao += 1

    # 3. Processamento com Pandas
    try:
        df = pd.read_excel(caminho_entrada)
        
        col_comp = 'complexidade' if 'complexidade' in df.columns else 'Complexidade'
        col_tipo = 'Tipo de desenvolvimento' if 'Tipo de desenvolvimento' in df.columns else 'Tipo de Desenvolvimento'
        
        if col_comp not in df.columns or col_tipo not in df.columns:
            messagebox.showerror("Erro", f"Colunas '{col_tipo}' ou '{col_comp}' não encontradas na planilha.")
            return

        def calcular(linha):
            t = str(linha.get(col_tipo, '')).strip().lower()
            c = str(linha.get(col_comp, '')).strip().lower()
            # Tenta achar a regra específica; se não achar, usa a regra Padrão (tipo='')
            return mapa_horas.get((t, c), mapa_horas.get(('', c), 0))

        df['Horas Estimadas'] = df.apply(calcular, axis=1)

        resumo = df.groupby(col_tipo, as_index=False).agg(
            Qtd_Tarefas=('ID', 'count'),
            Total_Horas=('Horas Estimadas', 'sum')
        )

        with pd.ExcelWriter(caminho_final, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Detalhamento', index=False)
            resumo.to_excel(writer, sheet_name='Resumo por Tipo', index=False)

        messagebox.showinfo("Sucesso!", f"Versão v{versao} gerada com sucesso!\nSalvo em: {pasta_destino}")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Falha no processamento:\n{str(e)}")

# ==========================================
# FUNÇÕES DA INTERFACE (DINÂMICA HIERÁRQUICA)
# ==========================================

lista_tipos = []

def adicionar_bloco_tipo(nome_tipo="", is_padrao=False):
    """Cria uma caixa (LabelFrame) para um Tipo de Desenvolvimento (Padrão ou Customizado)"""
    titulo = " ⭐ GRUPO PADRÃO (Regra Geral Obrigatória) " if is_padrao else " Grupo Específico "
    bloco_frame = ttk.LabelFrame(scrollable_frame, text=titulo, padding=10)
    bloco_frame.pack(fill="x", padx=10, pady=10)

    bloco_dict = {'frame_principal': bloco_frame, 'linhas': [], 'is_padrao': is_padrao}

    header_frame = ttk.Frame(bloco_frame)
    header_frame.pack(fill="x", pady=5)
    
    if is_padrao:
        ttk.Label(header_frame, text="Aplicado a qualquer Tipo de Desenvolvimento não mapeado abaixo.", font=("Arial", 9, "italic")).pack(side="left")
    else:
        ttk.Label(header_frame, text="Nome do Tipo:", font=("Arial", 9, "bold")).pack(side="left")
        e_tipo = ttk.Entry(header_frame, width=30)
        e_tipo.insert(0, nome_tipo)
        e_tipo.pack(side="left", padx=10)
        bloco_dict['entry_tipo'] = e_tipo
        # O botão Excluir só existe nos grupos que não são o padrão
        ttk.Button(header_frame, text="Excluir Grupo", command=lambda: remover_bloco(bloco_dict)).pack(side="right")

    linhas_frame = ttk.Frame(bloco_frame)
    linhas_frame.pack(fill="x", pady=5)
    bloco_dict['linhas_frame'] = linhas_frame

    # O botão de Adicionar nova complexidade só existe em grupos não-padrão
    if not is_padrao:
        ttk.Button(bloco_frame, text="+ Adicionar Complexidade", command=lambda: adicionar_linha_complexidade(bloco_dict)).pack(anchor="w", pady=5)

    lista_tipos.append(bloco_dict)
    return bloco_dict

def adicionar_linha_complexidade(bloco_dict, comp="", hora="", is_fixa=False):
    """Adiciona uma linha de [Complexidade | Horas | Botão X]"""
    row_frame = ttk.Frame(bloco_dict['linhas_frame'])
    row_frame.pack(fill="x", pady=2)

    ttk.Label(row_frame, text="↳ Complexidade:").pack(side="left", padx=(15, 5))
    e_comp = ttk.Entry(row_frame, width=20)
    e_comp.insert(0, comp)
    
    # Se a linha for fixa (como no Grupo Padrão), bloqueia a edição do nome e esconde o botão de remover
    if is_fixa:
        e_comp.config(state='readonly')
        
    e_comp.pack(side="left")

    ttk.Label(row_frame, text="Horas:").pack(side="left", padx=(15, 5))
    e_hora = ttk.Entry(row_frame, width=10)
    e_hora.insert(0, hora)
    e_hora.pack(side="left")

    linha_dict = {'frame': row_frame, 'comp': e_comp, 'hora': e_hora}
    
    if not is_fixa:
        tk.Button(row_frame, text="X", fg="red", relief="flat", cursor="hand2", command=lambda: remover_linha(bloco_dict, linha_dict)).pack(side="left", padx=10)

    bloco_dict['linhas'].append(linha_dict)

def remover_bloco(bloco_dict):
    bloco_dict['frame_principal'].destroy()
    lista_tipos.remove(bloco_dict)

def remover_linha(bloco_dict, linha_dict):
    linha_dict['frame'].destroy()
    bloco_dict['linhas'].remove(linha_dict)

# ==========================================
# UI PRINCIPAL
# ==========================================

app = tk.Tk()
app.title("Timesheet Automation Pro")
app.geometry("680x800")

style = ttk.Style()
style.theme_use('clam')

var_entrada = tk.StringVar(value=r"C:\Users\foca\Downloads\teste-script.xlsx")
var_saida = tk.StringVar(value=r"C:\Users\foca\OneDrive - Banco do Brasil S.A\Área de Trabalho\Timesheet\Cliente\Projeto")
var_nome_arquivo = tk.StringVar(value="planilha_saida_calculada_v")

# --- Seção 1: Ficheiros ---
f1 = ttk.LabelFrame(app, text=" 1. Caminhos ", padding=10)
f1.pack(fill="x", padx=20, pady=10)

ttk.Label(f1, text="Entrada:").grid(row=0, column=0, sticky="w")
ttk.Entry(f1, textvariable=var_entrada, width=60).grid(row=0, column=1, padx=5)
ttk.Button(f1, text="...", width=3, command=lambda: var_entrada.set(filedialog.askopenfilename())).grid(row=0, column=2)

ttk.Label(f1, text="Saída:").grid(row=1, column=0, sticky="w")
ttk.Entry(f1, textvariable=var_saida, width=60).grid(row=1, column=1, padx=5)
ttk.Button(f1, text="...", width=3, command=lambda: var_saida.set(filedialog.askdirectory())).grid(row=1, column=2)

ttk.Label(f1, text="Nome Base:").grid(row=2, column=0, sticky="w")
ttk.Entry(f1, textvariable=var_nome_arquivo, width=60).grid(row=2, column=1, padx=5, pady=5)

# --- Seção 2: Regras Dinâmicas (Área com Scroll) ---
f2 = ttk.LabelFrame(app, text=" 2. Regras de Esforço ", padding=5)
f2.pack(fill="both", expand=True, padx=20, pady=5)

canvas = tk.Canvas(f2, highlightthickness=0)
scrollbar = ttk.Scrollbar(f2, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=610)
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

btn_frame = ttk.Frame(app)
btn_frame.pack(fill="x", padx=20, pady=(5, 10))
ttk.Button(btn_frame, text="+ Adicionar Novo Tipo de Desenvolvimento Específico", command=adicionar_bloco_tipo).pack(fill="x")

# --- INICIALIZAÇÃO DE GRUPOS E REGRAS ---

# 1. Cria o Grupo Padrão (Fixo e Inalterável)
bloco_padrao = adicionar_bloco_tipo(is_padrao=True)
# Adiciona as 5 complexidades obrigatórias travadas (is_fixa=True)
for nivel in ["very low", "low", "medium", "high", "very high"]:
    adicionar_linha_complexidade(bloco_padrao, comp=nivel, is_fixa=True)

# 2. Cria um Grupo de Exemplo (Para mostrar como usar)
bloco_exemplo = adicionar_bloco_tipo("Front-end")
adicionar_linha_complexidade(bloco_exemplo, comp="low", hora="4")
adicionar_linha_complexidade(bloco_exemplo, comp="high", hora="12")

# --- Botão Final ---
tk.Button(app, text="PROCESSAR E GERAR PLANILHA", bg="#0052cc", fg="white", 
          font=("Arial", 12, "bold"), command=executar_processamento).pack(fill="x", padx=20, pady=10, ipady=10)

app.mainloop()