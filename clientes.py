# clientes.py
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os

# ---------------------- Helpers / Tratamento de dados ----------------------

def get_numbers_from_string(s):
    """Extrai apenas os dígitos de uma string. Retorna '' se vazio ou somente zeros."""
    try:
        s = str(s)
    except:
        return ''
    numbers = re.findall(r'\d+', s)
    texto = ''.join(numbers)
    if texto == "":
        return ""
    if all(caractere == '0' for caractere in texto):
        return ''
    return texto

def remover_itens_na_string(s, replacer_mask):
    """Substitui itens indesejados em uma string conforme replacer_mask dict."""
    try:
        s = str(s)
    except:
        return ''
    for item, rep in replacer_mask.items():
        s = s.replace(item, rep)
    return s

# máscara padrão sugerida (ajustável)
SUGGESTED_MASK = {
    'Nome/Razão Social': "nome",
    'CPF/CNPJ': 'cnpj_cpf',
    'Nome Fantasia': 'fantasia',
    'Inscrição Estadual': "ie",
    'Logradouro': "endereco",
    'Número': 'numero',
    'Complemento': 'complemento',
    'CEP': 'cep',
    'Bairro': 'bairro',
    'Cidade': 'cidade',
    'Estado': 'uf',
    'Telefone': "fone",
    'E-mail': 'email',
    'Funcionário': 'ncompr'
}

# colunas finais desejadas (ordem)
FINAL_COLUMNS = [
    'cliente_id', 'cnpj_cpf', 'cpf', 'ie', 'nome', 'fantasia', 'fone', 'endereco', 'numero',
    'bairro', 'cidade', 'ibge', 'uf', 'cep', 'email', 'email_danfe', 'classe_id', 'repr_id',
    'fone2', 'complemento', 'ncompr', 'rota_id', 'rota', 'consumidor_final', 'Observacao'
]

# ---------------------- Classe de processamento (sem GUI) ----------------------

class ClientesTratamento:
    """
    Classe com métodos de transformação dos dados de clientes.
    """

    def __init__(self,
                 numeric_fields=None,
                 replacer_mask=None,
                 uppercase_all=True):
        """
        numeric_fields: lista de colunas que serão tratadas com get_numbers_from_string
        replacer_mask: dicionário para remover itens indesejados em strings
        uppercase_all: se True, transforma colunas de texto em caixa alta
        """
        # conforme confirmação do usuário
        if numeric_fields is None:
            numeric_fields = ['cnpj_cpf','cpf','cep','fone','fone2','ie','cliente_id','ibge']
        self.numeric_fields = numeric_fields

        if replacer_mask is None:
            replacer_mask = {'S/N': '', 'SN': '', 'NAN': '', "'": ''}
        self.replacer_mask = replacer_mask

        self.uppercase_all = uppercase_all

    def clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Aplica limpeza geral: remove colunas vazias, normaliza textos, aplica remover_itens_na_string e get_numbers...
        Retorna dataframe reordenado com FINAL_COLUMNS.
        """
        # remover colunas totalmente vazias
        df = df.dropna(axis=1, how='all').copy()

        # transformar todos objetos em string e aplicar caixa alta se configurado
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str)
                if self.uppercase_all:
                    df[col] = df[col].str.upper()

        # remover itens indesejados em colunas texto
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = [remover_itens_na_string(s, self.replacer_mask) for s in df[col].values]

        # aplicar extração de números nas colunas configuradas
        for c in self.numeric_fields:
            if c in df.columns:
                df[c] = [get_numbers_from_string(x) for x in df[c].astype(str).values]

        # criar Observacao se existirem canais/ie/ponto_referencia (compat com notebook)
        if {'canal', 'ie', 'ponto_referencia'}.intersection(set(df.columns)):
            observacoes = []
            # preencher com '' caso alguma coluna não exista para evitar erro
            cols_to_use = [c for c in ['canal', 'ie', 'ponto_referencia'] if c in df.columns]
            df_fill = df[cols_to_use].fillna('')
            for row in df_fill.itertuples(index=False, name=None):
                obss = []
                if 'canal' in cols_to_use and row[cols_to_use.index('canal')] != "":
                    obss.append(f"Canal: {row[cols_to_use.index('canal')]}")
                if 'ie' in cols_to_use and row[cols_to_use.index('ie')] != "":
                    obss.append(f"IE/RG: {row[cols_to_use.index('ie')]}")
                if 'ponto_referencia' in cols_to_use and row[cols_to_use.index('ponto_referencia')] != "":
                    obss.append(f"Ponto de Referência: {row[cols_to_use.index('ponto_referencia')]}")
                observacoes.append("\n".join(obss) if obss else "")
            df['Observacao'] = observacoes

        # Garantir todas as colunas finais estão presentes
        for c in FINAL_COLUMNS:
            if c not in df.columns:
                df[c] = ''

        # Reordenar para a ordem desejada
        df = df[FINAL_COLUMNS]

        return df

# ---------------------- GUI para Clientes (3 etapas) ----------------------

class ClientesGUI:
    """
    Interface Tkinter para as 3 etapas:
    1) Selecionar arquivo
    2) Mapear/renomear colunas (ou deixar em branco para IGNORAR)
    3) Escolher tipos/formatações e executar processamento
    """

    def __init__(self, master=None, initial_filepath=None):
        self.master_owner = master
        # cria Toplevel se houver master, caso contrário cria root
        self.root = tk.Toplevel(master) if master else tk.Tk()
        self.root.title("Tratamento - Clientes")
        self.root.geometry("900x650")
        self.filepath = initial_filepath
        self.df_original = None
        self.df_mapped = None
        self.rename_map = {}
        self.ignored_columns = set()
        self.suggested_mask = SUGGESTED_MASK.copy()

        # variáveis de etapa 3
        self.numeric_vars = {}  # col -> tk.IntVar
        self.uppercase_var = tk.IntVar(value=1)
        # default replacer mask string (aparece no input)
        self.replacer_entry_var = tk.StringVar(value="S/N,SN,NAN,'")
        self.extra_numeric_var = tk.StringVar(value="")  # comma separated

        self._build_step1()

    # ---------------------- Step 1 ----------------------
    def _build_step1(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Etapa 1 — Selecionar arquivo de clientes", font=("Arial", 14)).pack(pady=6)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=8)

        ttk.Button(btn_frame, text="Selecionar arquivo", command=self._select_file).pack(side='left', padx=4)
        ttk.Button(btn_frame, text="Usar arquivo já carregado", command=self._use_loaded_file).pack(side='left', padx=4)

        self.file_label = ttk.Label(frame, text=self.filepath or "Nenhum arquivo selecionado")
        self.file_label.pack(pady=8)

        ttk.Label(frame, text="(Aceita .xlsx, .xls ou .csv)").pack()

        ttk.Button(frame, text="Avançar →", command=self._load_and_go_step2).pack(side='bottom', pady=12)

    def _select_file(self):
        path = filedialog.askopenfilename(title="Selecione a planilha de clientes",
                                          filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")])
        if path:
            self.filepath = path
            self.file_label.config(text=self.filepath)

    def _use_loaded_file(self):
        if not self.filepath:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
            return
        messagebox.showinfo("Info", f"Arquivo selecionado:\n{self.filepath}")

    def _load_and_go_step2(self):
        if not self.filepath:
            messagebox.showwarning("Aviso", "Selecione um arquivo antes de avançar.")
            return
        try:
            if self.filepath.lower().endswith('.csv'):
                df = pd.read_csv(self.filepath)
            else:
                df = pd.read_excel(self.filepath, skiprows=0)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo:\n{e}")
            return

        # drop empty columns
        df = df.dropna(axis=1, how='all')
        self.df_original = df
        self._build_step2()

    # ---------------------- Step 2 ----------------------
    def _build_step2(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        ttk.Label(self.root, text="Etapa 2 — Mapear colunas (deixe em branco para IGNORAR)", font=("Arial", 13)).pack(pady=6)

        container = ttk.Frame(self.root)
        container.pack(fill='both', expand=True, padx=8, pady=8)

        # Scrollable canvas for many columns
        canvas = tk.Canvas(container)
        vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=380)

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self.column_entries = {}  # original_col -> Entry widget

        # header
        hdr = ttk.Frame(scroll_frame)
        hdr.grid(row=0, column=0, sticky='ew', pady=2)
        ttk.Label(hdr, text="Coluna original", width=50).grid(row=0, column=0, padx=2)
        ttk.Label(hdr, text="Nome destino (deixe em branco para IGNORAR)", width=50).grid(row=0, column=1, padx=2)

        for idx, col in enumerate(self.df_original.columns, start=1):
            lbl = ttk.Label(scroll_frame, text=str(col), width=50, anchor='w')
            lbl.grid(row=idx, column=0, padx=2, pady=2, sticky='w')

            # sugestão de rename: se estiver em SUGGESTED_MASK, usar; senão pegar nome simplificado
            sugest = self.suggested_mask.get(col, col)
            ent = ttk.Entry(scroll_frame, width=50)
            ent.insert(0, sugest)
            ent.grid(row=idx, column=1, padx=2, pady=2, sticky='w')
            self.column_entries[col] = ent

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill='x', pady=8)
        ttk.Button(btn_frame, text="← Voltar", command=self._build_step1).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Avançar →", command=self._process_step2).pack(side='right', padx=6)

    def _process_step2(self):
        # Build rename_map and ignored list: agora campo em branco = ignorar
        rename_map = {}
        ignored = set()
        for orig, ent in self.column_entries.items():
            val = ent.get().strip()
            if val == "":
                ignored.add(orig)
            else:
                rename_map[orig] = val

        # aplicar renome temporariamente
        df = self.df_original.rename(columns=rename_map).copy()

        # remover colunas ignoradas (usando orig names)
        df.drop(columns=list(ignored), inplace=True, errors='ignore')

        self.df_mapped = df
        self.rename_map = rename_map
        self.ignored_columns = ignored

        self._build_step3()

    # ---------------------- Step 3 ----------------------
    def _build_step3(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        ttk.Label(self.root, text="Etapa 3 — Definir tipos / formatos", font=("Arial", 13)).pack(pady=6)

        frame = ttk.Frame(self.root)
        frame.pack(fill='both', expand=True, padx=8, pady=8)

        # Options: uppercase all
        opts = ttk.Frame(frame)
        opts.pack(anchor='w', pady=6)
        ttk.Checkbutton(opts, text="Transformar textos para CAIXA ALTA", variable=self.uppercase_var).pack(side='left', padx=6)

        ttk.Label(opts, text="Remover itens (separe por vírgula):").pack(side='left', padx=6)
        ttk.Entry(opts, textvariable=self.replacer_entry_var, width=30).pack(side='left', padx=6)

        # extra numeric input
        ttk.Label(frame, text="Colunas adicionais para extrair apenas números (separe por vírgula):").pack(anchor='w', pady=4)
        ttk.Entry(frame, textvariable=self.extra_numeric_var, width=60).pack(anchor='w')

        ttk.Label(frame, text="Marque as colunas que devem ficar apenas com números:", font=("Arial", 11)).pack(anchor='w', pady=8)

        cols_container = ttk.Frame(frame)
        cols_container.pack(fill='both', expand=True)

        # Scrollable list of columns with checkboxes
        canvas = tk.Canvas(cols_container)
        vsb = ttk.Scrollbar(cols_container, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=vsb.set, height=300)

        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        self.numeric_vars = {}
        for idx, col in enumerate(self.df_mapped.columns):
            # padrão: marcar como numéricas as colunas confirmadas
            initial = 1 if col.lower() in ['cnpj_cpf','cpf','cep','fone','fone2','ie','cliente_id','ibge'] else 0
            var = tk.IntVar(value=initial)
            cb = ttk.Checkbutton(inner, text=col, variable=var)
            cb.grid(row=idx, column=0, sticky='w', padx=6, pady=2)
            self.numeric_vars[col] = var

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill='x', pady=8)
        ttk.Button(btn_frame, text="← Voltar", command=self._build_step2).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Visualizar preview", command=self._show_preview).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Processar e Salvar", command=self._process_and_save).pack(side='right', padx=6)

    def _show_preview(self):
        tratador = self._build_tratador_from_ui()
        try:
            df_preview = tratador.clean_dataframe(self.df_mapped.copy())
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar preview:\n{e}")
            return
        PreviewWindow(self.root, df_preview.head(200))

    def _build_tratador_from_ui(self):
        # get numeric selection
        numeric_selected = [c for c, var in self.numeric_vars.items() if var.get() == 1]
        extra_txt = self.extra_numeric_var.get().strip()
        extra_list = [s.strip() for s in extra_txt.split(',') if s.strip()] if extra_txt else []
        # build replacer mask from entry
        replacer_input = self.replacer_entry_var.get().strip()
        replacer_mask = {}
        if replacer_input:
            parts = [p.strip() for p in replacer_input.split(',') if p.strip()]
            for p in parts:
                replacer_mask[p] = ''

        # combine numeric_selected with extra_list (unique)
        numeric_final = list(dict.fromkeys(numeric_selected + extra_list))

        tratador = ClientesTratamento(
            numeric_fields=numeric_final,
            replacer_mask=replacer_mask,
            uppercase_all=bool(self.uppercase_var.get())
        )
        return tratador

    def _process_and_save(self):
        tratador = self._build_tratador_from_ui()
        try:
            df_final = tratador.clean_dataframe(self.df_mapped.copy())
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar dados:\n{e}")
            return

        # Pergunta onde salvar
        default_dir = os.path.dirname(self.filepath) if self.filepath else os.getcwd()
        suggested_name = os.path.join(default_dir, os.path.basename(os.path.splitext(self.filepath)[0]) + " - Clientes Tratado.xlsx")
        out_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                initialfile=os.path.basename(suggested_name),
                                                initialdir=default_dir,
                                                filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not out_path:
            return
        try:
            df_final.to_excel(out_path, index=False)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo:\n{e}")
            return

        messagebox.showinfo("Concluído", f"Arquivo salvo em:\n{out_path}")
        # fechar janela se for Toplevel
        if isinstance(self.root, tk.Toplevel):
            self.root.destroy()

    # ---------------------- start / helper ----------------------
    def start(self):
        # iniciar mainloop se esta janela é a root
        if isinstance(self.root, tk.Tk):
            self.root.mainloop()
        else:
            # se é Toplevel, apenas retorna para que o caller mantenha o mainloop
            return

# ---------------------- Janela de Preview (Treeview com scrollbars) ----------------------

class PreviewWindow:
    def __init__(self, master, df_preview: pd.DataFrame):
        self.win = tk.Toplevel(master)
        self.win.title("Preview - primeiras linhas")
        # tamanho ajustável; usuário pode redimensionar
        self.win.geometry("1000x500")

        frame = ttk.Frame(self.win, padding=6)
        frame.pack(fill='both', expand=True)

        cols = list(df_preview.columns)

        # Treeview com scrollbars configuradas (vertical + horizontal)
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill='both', expand=True)

        # vertical and horizontal scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        tree = ttk.Treeview(tree_frame, columns=cols, show='headings',
                            yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # configurar scrollbars para controlar tree
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)

        # layout grid para tree + scrollbars
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # cabeçalhos e largura inicial
        for c in cols:
            tree.heading(c, text=c)
            # ajustar width dinamicamente (padrão 140)
            tree.column(c, width=140, anchor='w')

        # inserir linhas (convertendo tudo para string)
        for row in df_preview.itertuples(index=False):
            tree.insert('', 'end', values=[str(x) for x in row])

        # permitir redimensionar coluna ao clicar no separador de header (comportamento padrão)

        # barra de fechamento
        btns = ttk.Frame(frame)
        btns.pack(fill='x', pady=6)
        ttk.Button(btns, text="Fechar", command=self.win.destroy).pack(side='right', padx=6)

# ---------------------- Módulo de teste ----------------------
if __name__ == "__main__":
    gui = ClientesGUI()
    gui.start()