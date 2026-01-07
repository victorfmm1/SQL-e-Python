# fornecedores.py
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os

# ---------------------- Helpers ----------------------

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

def remover_itens_na_string(s, replacer_mask):
    try:
        s = str(s)
    except:
        return ""
    for item, rep in replacer_mask.items():
        s = s.replace(item, rep)
    return s


# ---------------------- MÁSCARA PADRÃO ----------------------

DEFAULT_REPLACER_MASK = {
    "S/N": "",
    "SN": "",
    "NAN": "",
    "'": ""
}

FINAL_COLUMNS_FORNECEDORES = [
    "fornecedor_id", "cnpj_cpf", "ie", "razao", "fantasia", "fone", "fone2", "email",
    "email_nfe", "endereco", "numero", "complemento", "bairro",
    "cidade", "uf", "cep", "ibge", "contato", "observacao"
]


# ---------------------- Classe de Tratamento ----------------------

class FornecedoresTratamento:

    def __init__(self,
                 numeric_fields=None,
                 replacer_mask=None,
                 uppercase_all=True):

        if numeric_fields is None:
            numeric_fields = [
                "cnpj_cpf", "cep", "fone", "fone2", "ie", "ibge", "fornecedor_id"
            ]

        self.numeric_fields = numeric_fields
        self.replacer_mask = replacer_mask or DEFAULT_REPLACER_MASK.copy()
        self.uppercase_all = uppercase_all

    def clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:

        df = df.dropna(axis=1, how="all").copy()

        # padronizar texto
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str)
                if self.uppercase_all:
                    df[col] = df[col].str.upper()
                df[col] = [remover_itens_na_string(x, self.replacer_mask) for x in df[col]]

        # aplicar extração numérica
        for col in self.numeric_fields:
            if col in df.columns:
                df[col] = [get_numbers_from_string(x) for x in df[col].astype(str)]

        # gerar OBSERVAÇÃO final a partir de campos importantes
        obs_list = []
        for _, row in df.fillna("").iterrows():
            temp = []

            if "contato" in df.columns and row["contato"].strip():
                temp.append(f"Contato: {row['contato']}")

            if "ie" in df.columns and row["ie"].strip():
                temp.append(f"IE/RG: {row['ie']}")

            obs_list.append("\n".join(temp) if temp else "")

        df["observacao"] = obs_list

        # garantir todas as colunas finais
        for col in FINAL_COLUMNS_FORNECEDORES:
            if col not in df.columns:
                df[col] = ""

        return df[FINAL_COLUMNS_FORNECEDORES]


# ---------------------- GUI (3 passos) ----------------------

class FornecedoresGUI:

    def __init__(self, master=None, initial_filepath=None):
        self.master_owner = master
        self.root = tk.Toplevel(master) if master else tk.Tk()
        self.root.title("Tratamento - Fornecedores")
        self.root.geometry("900x650")

        self.filepath = initial_filepath
        self.df_original = None
        self.df_mapped = None

        # mapeamento das colunas
        self.rename_map = {}

        self.numeric_vars = {}
        self.uppercase_var = tk.IntVar(value=1)

        self.replacer_entry_var = tk.StringVar(value="S/N,SN,NAN,'")
        self.extra_numeric_var = tk.StringVar(value="")

        self._build_step1()

    # ---------------- STEP 1 ----------------
    def _build_step1(self):
        for w in self.root.winfo_children():
            w.destroy()

        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Etapa 1 — Selecionar arquivo de fornecedores",
                  font=("Arial", 14)).pack(pady=6)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Selecionar arquivo",
                   command=self._select_file).pack(side="left", padx=4)

        ttk.Button(btn_frame, text="Usar arquivo já carregado",
                   command=self._use_loaded_file).pack(side="left", padx=4)

        self.file_label = ttk.Label(frame, text=self.filepath or "Nenhum arquivo selecionado")
        self.file_label.pack(pady=8)

        ttk.Button(frame, text="Avançar →", command=self._load_and_go_step2).pack(pady=12)

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")]
        )
        if path:
            self.filepath = path
            self.file_label.config(text=path)

    def _use_loaded_file(self):
        if not self.filepath:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
        else:
            messagebox.showinfo("Arquivo", self.filepath)

    def _load_and_go_step2(self):
        if not self.filepath:
            messagebox.showwarning("Aviso", "Selecione um arquivo.")
            return

        try:
            if self.filepath.lower().endswith(".csv"):
                df = pd.read_csv(self.filepath)
            else:
                df = pd.read_excel(self.filepath)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler arquivo:\n{e}")
            return

        df = df.dropna(axis=1, how="all")
        self.df_original = df
        self._build_step2()

    # ---------------- STEP 2 ----------------
    def _build_step2(self):
        for w in self.root.winfo_children():
            w.destroy()

        ttk.Label(self.root, text="Etapa 2 — Mapear colunas",
                  font=("Arial", 14)).pack(pady=6)

        container = ttk.Frame(self.root)
        container.pack(fill="both", expand=True, padx=8, pady=8)

        canvas = tk.Canvas(container)
        vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        inner = ttk.Frame(canvas)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=380)

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        # construir widgets
        self.column_entries = {}

        for idx, col in enumerate(self.df_original.columns):
            ttk.Label(inner, text=f"{col}").grid(row=idx, column=0, padx=5, pady=3, sticky="w")

            entry = ttk.Entry(inner, width=30)
            entry.grid(row=idx, column=1, padx=5, pady=3)

            # se tiver correspondência automática
            for final_col in FINAL_COLUMNS_FORNECEDORES:
                if final_col.lower() in col.lower():
                    entry.insert(0, final_col)

            self.column_entries[col] = entry

        ttk.Button(self.root, text="Avançar →", command=self._go_step3).pack(pady=12)

    def _go_step3(self):
        mapping = {}
        for orig, entry in self.column_entries.items():
            new_name = entry.get().strip()
            mapping[orig] = new_name if new_name else None

        self.rename_map = mapping
        self._build_step3()

    # ---------------- STEP 3 ----------------
    def _build_step3(self):
        for w in self.root.winfo_children():
            w.destroy()

        ttk.Label(self.root,
                  text="Etapa 3 — Configurações finais",
                  font=("Arial", 14)).pack(pady=8)

        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Checkbutton(frame, text="Transformar textos em CAIXA ALTA",
                        variable=self.uppercase_var).pack(anchor="w", pady=4)

        ttk.Label(frame, text="Máscara para remover itens (separar por vírgula):").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.replacer_entry_var,
                  width=50).pack(anchor="w", pady=4)

        ttk.Label(frame, text="Forçar extração numérica em colunas (separado por vírgula):").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.extra_numeric_var,
                  width=50).pack(anchor="w", pady=4)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Gerar preview",
                   command=self._preview).pack(side="left", padx=10)

        ttk.Button(btn_frame, text="Salvar arquivo final",
                   command=self._save).pack(side="left", padx=10)

    def _apply_mapping(self):
        df = self.df_original.copy()

        new_cols = {}
        ignored = []

        for orig, new in self.rename_map.items():
            if new is None or new == "":
                ignored.append(orig)
            else:
                new_cols[orig] = new

        if len(new_cols) == 0:
            messagebox.showwarning(
                "Aviso",
                "Nenhuma coluna foi mapeada. Mapear ao menos uma coluna."
            )
            return None

        df = df.rename(columns=new_cols)

        # drop seguro
        if ignored:
            df = df.drop(columns=ignored, errors="ignore")

        return df

    def _preview(self):
        df = self._apply_mapping()
        if df is None or df.empty:
            messagebox.showwarning("Aviso", "Nada a processar após o mapeamento.")
            return

        df = df.dropna(axis=1, how="all")

        replacer = {
            x.strip(): ""
            for x in self.replacer_entry_var.get().split(",")
            if x.strip()
        }

        # extra numeric → manter padrão se vazio
        extra_raw = [
            x.strip()
            for x in self.extra_numeric_var.get().split(",")
            if x.strip()
        ]
        extra_numeric = extra_raw if extra_raw else None

        try:
            tratamento = FornecedoresTratamento(
                numeric_fields=extra_numeric,
                replacer_mask=replacer,
                uppercase_all=bool(self.uppercase_var.get())
            )

            df_final = tratamento.clean_dataframe(df)

        except Exception as e:
            messagebox.showerror("Erro no Preview", f"Falha ao processar dados:\n{e}")
            return

        # janela de preview
        win = tk.Toplevel(self.root)
        win.title("Preview")
        win.geometry("900x500")

        tree = ttk.Treeview(win, columns=list(df_final.columns), show="headings")
        tree.pack(fill="both", expand=True)

        for col in df_final.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120)

        for _, row in df_final.head(50).iterrows():
            tree.insert("", "end", values=list(row.values))

        ttk.Button(win, text="Fechar", command=win.destroy).pack(pady=10)

    def _save(self):
        df = self._apply_mapping()
        if df is None or df.empty:
            messagebox.showwarning("Aviso", "Nada a processar após o mapeamento.")
            return

        df = df.dropna(axis=1, how="all")

        replacer = {
            x.strip(): ""
            for x in self.replacer_entry_var.get().split(",")
            if x.strip()
        }

        # extra numeric → manter padrão se vazio
        extra_raw = [
            x.strip()
            for x in self.extra_numeric_var.get().split(",")
            if x.strip()
        ]
        extra_numeric = extra_raw if extra_raw else None

        try:
            tratamento = FornecedoresTratamento(
                numeric_fields=extra_numeric,
                replacer_mask=replacer,
                uppercase_all=bool(self.uppercase_var.get())
            )

            df_final = tratamento.clean_dataframe(df)

        except Exception as e:
            messagebox.showerror("Erro ao salvar", f"Falha ao processar dados:\n{e}")
            return

        out_path = os.path.join(
            os.path.dirname(self.filepath),
            "fornecedores_tratado.xlsx"
        )

        try:
            df_final.to_excel(out_path, index=False)
        except Exception as e:
            messagebox.showerror("Erro ao salvar arquivo", f"Não foi possível salvar:\n{e}")
            return

        messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{out_path}")


if __name__ == "__main__":
    gui = FornecedoresGUI()
    gui.root.mainloop()
