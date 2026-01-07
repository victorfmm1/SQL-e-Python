# produtos.py
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os

DEFAULT_REPLACER_MASK = {'S/N': '', 'SN': '', 'NAN': '', "'": ''}
DEFAULT_NUMERIC = ['ncm', 'origem', 'ean13', 'cest']

# ---------------------- Helpers / tratamento ----------------------

def get_numbers_from_string(s):
    try:
        s = str(s)
    except:
        return ""
    nums = re.findall(r"\d+", s)
    if not nums:
        return ""
    texto = "".join(nums)
    if all(ch == "0" for ch in texto):
        return ""
    return texto

def remove_items_in_string(s, replacer_mask):
    try:
        s = str(s)
    except:
        return ""
    for k, v in replacer_mask.items():
        s = s.replace(k, v)
    return s

def normalize_decimal_to_comma(v):
    # transforma pontos decimais em vírgula (mantém strings vazias)
    try:
        s = str(v)
    except:
        return ""
    if s.strip() == "":
        return ""
    return s.replace(".", ",")

def clean_decimal_value(s):
    try:
        s = str(s).strip().upper()
    except:
        return "0"

    if s == "" or s in ["NAN", "NONE"]:
        return "0"

    # Remover símbolos
    s = re.sub(r"[^\d,.-]", "", s)

    # Se tiver vírgula e ponto: remover separador de milhares
    if "," in s and "." in s:
        # Ex: 1.234,56 → remover ponto
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
        else:
            s = s.replace(",", "")

    # Agora padronizar: trocar vírgula por ponto
    s = s.replace(",", ".")

    try:
        val = float(s)
        return f"{val:.2f}".replace(".", ",")  # volta para vírgula
    except:
        return "0"


# ---------------------- Máscara e colunas finais (ADSNet) ----------------------

MASCARA_FIXA = {
    'Nome do Produto (120)': 'descricao',
    'Fornecedor': 'fornece',
    'Marca (25)': 'marca',
    'Estoque Atual': 'estoque',
    'Unidade (06)': 'unidade',
    'Valor Venda (Tabela Padrão)': 'prvenda',
    'Valor Custo': 'ccompra',
    'Peso': 'pesobr',
    'Código CEST': 'cest',
    'NCM (8)': 'ncm',
    'Origem (0 a 8)': 'origem',
    'Código de Barras (GTIN-8,12,13,14)': 'ean13'
}

FINAL_COLUMNS_ADSNET = [
    "produto_id", "descricao", "apresentacao", "marca", "codfab", "ean13",
    "ncm", "cest", "estoque", "unidade", "fator_cv", "muv", "localizacao",
    "descricao2", "divisao_id", "fornece", "fornec_id", "trib_icms", "cst",
    "csosn", "aliq_icms", "pFCP", "pauta", "cicms", "cIpi", "cst_pis",
    "cst_cofins", "origem", "pesobr", "u_nota", "ccompra", "cmedio",
    "prvenda", "comissao", "unidade2", "dun14", "pno"
]

# ---------------------- Classe de tratamento (sem GUI) ----------------------

class ProdutosTratamento:
    def __init__(self, numeric_fields=None, decimal_fields=None,replacer_mask=None, uppercase_all=True):
        self.numeric_fields = numeric_fields or DEFAULT_NUMERIC.copy()
        self.decimal_fields = decimal_fields or []
        self.replacer_mask = replacer_mask or DEFAULT_REPLACER_MASK.copy()
        self.uppercase_all = uppercase_all

    def clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.dropna(axis=1, how='all').copy()

        # transformar textos e aplicar uppercase se solicitado
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str)
                if self.uppercase_all:
                    df[col] = df[col].str.upper()
                df[col] = [remove_items_in_string(x, self.replacer_mask) for x in df[col].values]

        # tratar decimais (preços/pesos) se existirem
        for col in self.decimal_fields:
            if col in df.columns:
                df[col] = [clean_decimal_value(x) for x in df[col].astype(str)]

        # aplicar extração numérica nas colunas configuradas
        for c in self.numeric_fields:
            if c in df.columns:
                df[c] = [get_numbers_from_string(x) for x in df[c].astype(str).values]

        # remover trailing ".0" que costumam aparecer em colunas de códigos
        for c in ['ean13', 'ncm', 'cest']:
            if c in df.columns:
                df[c] = df[c].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # garantir todas as colunas ADSNet existam (preencher vazias)
        for c in FINAL_COLUMNS_ADSNET:
            if c not in df.columns:
                df[c] = ''

        # reordenar e retornar
        return df[FINAL_COLUMNS_ADSNET]

# ---------------------- Preview Window (Treeview com scroll) ----------------------

class PreviewWindow:
    def __init__(self, master, df_preview: pd.DataFrame, title="Preview"):
        self.win = tk.Toplevel(master)
        self.win.title(title)
        self.win.geometry("1000x520")

        frame = ttk.Frame(self.win, padding=6)
        frame.pack(fill='both', expand=True)

        cols = list(df_preview.columns)

        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill='both', expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings',
                                 yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, anchor='w')

        for row in df_preview.itertuples(index=False):
            self.tree.insert('', 'end', values=[str(x) for x in row])

        btns = ttk.Frame(frame)
        btns.pack(fill='x', pady=6)
        ttk.Button(btns, text="Fechar", command=self.win.destroy).pack(side='right', padx=6)

# ---------------------- GUI Produtos (3 etapas) ----------------------

class ProdutosGUI:
    """
    Fluxo:
    1) Selecionar arquivo
    2) Mapear renomear colunas (deixe em branco para ignorar)
    3) Configurações finais (uppercase, remover itens, extra numeric) + preview + salvar
    """

    def __init__(self, master=None, initial_filepath=None):
        self.master_owner = master
        self.root = tk.Toplevel(master) if master else tk.Tk()
        self.root.title("Tratamento - Produtos")
        self.root.geometry("900x650")

        self.filepath = initial_filepath
        self.df_original = None
        self.df_mapped = None

        self.rename_map = {}
        self.ignored_columns = set()

        self.numeric_vars = {}
        self.uppercase_var = tk.IntVar(value=1)
        self.replacer_entry_var = tk.StringVar(value="S/N,SN,NAN,'")
        self.extra_numeric_var = tk.StringVar(value="")  # cols comma separated

        self._build_step1()

    # ---- Step 1: escolher arquivo ----
    def _build_step1(self):
        for w in self.root.winfo_children():
            w.destroy()

        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Etapa 1 — Selecionar arquivo de produtos", font=("Arial", 14)).pack(pady=6)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=8)

        ttk.Button(btn_frame, text="Selecionar arquivo", command=self._select_file).pack(side='left', padx=4)
        ttk.Button(btn_frame, text="Usar arquivo já carregado", command=self._use_loaded_file).pack(side='left', padx=4)

        self.file_label = ttk.Label(frame, text=self.filepath or "Nenhum arquivo selecionado")
        self.file_label.pack(pady=8)
        ttk.Label(frame, text="(Aceita .xlsx, .xls ou .csv)").pack()

        ttk.Button(frame, text="Avançar →", command=self._load_and_go_step2).pack(side='bottom', pady=12)

    def _select_file(self):
        path = filedialog.askopenfilename(title="Selecione a planilha de produtos",
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

        df = df.dropna(axis=1, how='all')
        self.df_original = df
        self._build_step2()

    # ---- Step 2: mapear colunas ----
    def _build_step2(self):
        for w in self.root.winfo_children():
            w.destroy()

        ttk.Label(self.root, text="Etapa 2 — Mapear colunas (deixe em branco para IGNORAR)", font=("Arial", 13)).pack(pady=6)

        container = ttk.Frame(self.root)
        container.pack(fill='both', expand=True, padx=8, pady=8)

        # canvas scroll
        canvas = tk.Canvas(container)
        vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        inner = ttk.Frame(canvas)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=380)

        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        self.column_entries = {}

        # header
        hdr = ttk.Frame(inner)
        hdr.grid(row=0, column=0, sticky='ew')
        ttk.Label(hdr, text="Coluna original", width=50).grid(row=0, column=0, padx=2)
        ttk.Label(hdr, text="Nome destino (deixe em branco para IGNORAR)", width=50).grid(row=0, column=1, padx=2)

        for idx, col in enumerate(self.df_original.columns, start=1):
            ttk.Label(inner, text=str(col), width=50, anchor='w').grid(row=idx, column=0, padx=2, pady=2, sticky='w')
            ent = ttk.Entry(inner, width=50)
            # sugestão: se o nome original corresponder a alguma máscara fixa, já preenche
            sugest = MASCARA_FIXA.get(col, "")
            if sugest:
                ent.insert(0, sugest)
            self.column_entries[col] = ent
            ent.grid(row=idx, column=1, padx=2, pady=2, sticky='w')

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill='x', pady=8)
        ttk.Button(btn_frame, text="← Voltar", command=self._build_step1).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Avançar →", command=self._process_step2).pack(side='right', padx=6)

    def _process_step2(self):
        rename_map = {}
        ignored = set()
        for orig, ent in self.column_entries.items():
            val = ent.get().strip()
            if val == "":
                ignored.add(orig)
            else:
                rename_map[orig] = val

        df = self.df_original.rename(columns=rename_map).copy()
        df.drop(columns=list(ignored), inplace=True, errors='ignore')

        self.df_mapped = df
        self.rename_map = rename_map
        self.ignored_columns = ignored

        self._build_step3()

    # ---- Step 3: configurações e salvar ----
    def _build_step3(self):
        for w in self.root.winfo_children():
            w.destroy()

        ttk.Label(self.root, text="Etapa 3 — Definir tipos / formatos", font=("Arial", 13)).pack(pady=6)

        frame = ttk.Frame(self.root)
        frame.pack(fill='both', expand=True, padx=8, pady=8)

        opts = ttk.Frame(frame)
        opts.pack(anchor='w', pady=6)
        ttk.Checkbutton(opts, text="Transformar textos para CAIXA ALTA", variable=self.uppercase_var).pack(side='left', padx=6)
        ttk.Label(opts, text="Remover itens (separe por vírgula):").pack(side='left', padx=6)
        ttk.Entry(opts, textvariable=self.replacer_entry_var, width=30).pack(side='left', padx=6)

        ttk.Label(frame, text="Colunas adicionais para extrair apenas números (separe por vírgula):").pack(anchor='w', pady=4)
        ttk.Entry(frame, textvariable=self.extra_numeric_var, width=60).pack(anchor='w')

        ttk.Label(frame, text="Colunas com valores decimais (separar por vírgula):").pack(anchor="w")
        self.decimal_columns_var = tk.StringVar(value="")
        ttk.Entry(frame, textvariable=self.decimal_columns_var,width=50).pack(anchor="w", pady=4)

        ttk.Label(frame, text="Marque as colunas que devem ficar apenas com números:", font=("Arial", 11)).pack(anchor='w', pady=8)
        cols_container = ttk.Frame(frame)
        cols_container.pack(fill='both', expand=True)

        canvas = tk.Canvas(cols_container)
        vsb = ttk.Scrollbar(cols_container, orient='vertical', command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=vsb.set, height=300)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        self.numeric_vars = {}
        for idx, col in enumerate(self.df_mapped.columns):
            initial = 1 if col.lower() in DEFAULT_NUMERIC else 0
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
        PreviewWindow(self.root, df_preview.head(200), title="Preview - Produtos (primeiras linhas)")

    def _build_tratador_from_ui(self):
        numeric_selected = [c for c, var in self.numeric_vars.items() if var.get() == 1]
        extra_txt = self.extra_numeric_var.get().strip()
        extra_list = [s.strip() for s in extra_txt.split(',') if s.strip()] if extra_txt else []
        replacer_input = self.replacer_entry_var.get().strip()
        replacer_mask = {}
        if replacer_input:
            parts = [p.strip() for p in replacer_input.split(',') if p.strip()]
            for p in parts:
                replacer_mask[p] = ''
            
        decimal_txt = self.decimal_columns_var.get().strip()
        decimal_list = [s.strip() for s in decimal_txt.split(',') if s.strip()] if decimal_txt else []

        numeric_final = list(dict.fromkeys(numeric_selected + extra_list))
        if decimal_list:
            numeric_final = [c for c in numeric_final if c not in decimal_list]

        tratador = ProdutosTratamento(
            numeric_fields=numeric_final,
            decimal_fields=decimal_list,
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

        default_dir = os.path.dirname(self.filepath) if self.filepath else os.getcwd()
        suggested_name = os.path.join(default_dir, os.path.basename(os.path.splitext(self.filepath)[0]) + " - Produtos Tratado.xlsx")
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
        if isinstance(self.root, tk.Toplevel):
            self.root.destroy()

    # start/helper
    def start(self):
        if isinstance(self.root, tk.Tk):
            self.root.mainloop()
        else:
            return

# ---------------------- módulo de teste ----------------------

if __name__ == "__main__":
    gui = ProdutosGUI()
    gui.start()