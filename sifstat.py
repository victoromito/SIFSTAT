"""
SIFSTAT - Sistema Unificado de Conferência de Mapas Estatísticos JBS
=====================================================================
Unifica os três conferidores (Produção, Recebimento e Comercialização)
em um único programa com menu lateral de navegação.

Desenvolvido por:
Víctor Sanches - FATURAMENTO JBS-GYN
Heitor Lopes - CSC FISCAL JBS-GYN
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import unicodedata
import os
import sys
import csv
import re
import webbrowser

try:
    from PIL import Image, ImageTk
    _PIL_OK = True
except Exception:
    _PIL_OK = False

try:
    _RESAMPLE = Image.Resampling.LANCZOS
except Exception:
    try:
        _RESAMPLE = Image.LANCZOS
    except Exception:
        _RESAMPLE = None

# ==========================================================================
# CAMINHO DE RECURSOS (funciona rodando o .py e também depois de gerado .exe)
# ==========================================================================

def resource_path(relative_path):
    """Retorna o caminho absoluto de um recurso (imagem/ícone), tanto em modo
    script quanto quando compilado com PyInstaller (--onefile)."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


TOLERANCIA_KG = 0.0001

# ==========================================================================
# FUNÇÕES COMPARTILHADAS
# (idênticas nos três programas originais - mantidas em um único lugar)
# ==========================================================================

def format_br(num):
    return f"{num:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")


def normalize_header(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().lower()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return s


def normalize_product(val):
    if pd.isna(val):
        return ""
    s = str(val).replace('\u00A0', ' ')
    return " ".join(s.split()).upper()


def find_column(df, keyword_groups, max_rows=10):
    melhor_score = 0
    melhor_pos = (None, None)
    max_scan = min(max_rows, len(df))
    for r in range(max_scan):
        row = df.iloc[r]
        for c, val in enumerate(row):
            header = normalize_header(val)
            score = 0
            for grupo in keyword_groups:
                for palavra in grupo:
                    if header == palavra:
                        score += 100
                    elif palavra in header:
                        score += 25
            if score > melhor_score:
                melhor_score = score
                melhor_pos = (r, c)
    return melhor_pos


def remove_acentos(texto):
    if pd.isna(texto):
        return ""
    s = str(texto).strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def limpar_numero_simples(val):
    """Versão usada pelo módulo de Produção (sem limpeza por regex)."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return round(float(val), 4)
    s = str(val).strip()
    if not s:
        return 0.0
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return round(float(s), 4)
    except ValueError:
        return 0.0


def limpar_numero_regex(val):
    """Versão usada pelos módulos de Recebimento e Comercialização
    (remove qualquer caractere que não seja dígito, ponto ou vírgula)."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return round(float(val), 4)
    s = str(val).strip()
    if not s:
        return 0.0
    s = re.sub(r'[^\d.,]', '', s)
    if not s:
        return 0.0
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return round(float(s), 4)
    except ValueError:
        return 0.0


def normalize_sif(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().split('.')[0]
    s = ''.join(filter(str.isdigit, s))
    return s if s else "SEM_SIF"


def normalize_product_acentos(val):
    """Versão usada pelo módulo de Comercialização (remove acentos também)."""
    if pd.isna(val):
        return ""
    s = str(val).replace('\u00A0', ' ')
    return remove_acentos(" ".join(s.split()).upper())


def normalize_location(val, dicionario_ativo):
    if pd.isna(val):
        return ""
    loc = remove_acentos(str(val).upper())
    return dicionario_ativo.get(loc, loc)


# ==========================================================================
# MÓDULO: PRODUÇÃO
# ==========================================================================

class ProducaoFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="black")
        self.controller = controller
        self.erp_path = None
        self.pga_path = None
        self.last_results = []
        self.warnings_list = []
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="SIFStat (Produção)", font=("Segoe UI", 18, "bold"), bg="black", fg="white").pack(pady=(20, 5))
        tk.Label(self, text="Faça o upload do ERP ordenado e do PGA para conciliar os pesos.", font=("Segoe UI", 10), bg="black", fg="#cbd5e1").pack(pady=(0, 15))

        self.somente_divergencias = tk.BooleanVar(value=False)
        tk.Checkbutton(
            self,
            text="Mostrar apenas divergências",
            variable=self.somente_divergencias,
            bg="black", fg="white", selectcolor="black",
            activebackground="black", activeforeground="white",
            font=("Segoe UI", 10),
            command=self.atualizar_tabela
        ).pack(pady=(0, 10))

        frame_upload = tk.Frame(self, bg="black")
        frame_upload.pack(fill="x", pady=10, padx=20)

        self.btn_erp = tk.Button(frame_upload, text="📁 1. Selecionar ERP Ordenado", command=self.carregar_erp, width=40, height=2, bg="#ffffff", font=("Segoe UI", 10, "bold"), relief="ridge")
        self.btn_erp.pack(side="left", padx=10, expand=True)

        self.btn_pga = tk.Button(frame_upload, text="📁 2. Selecionar PGA", command=self.carregar_pga, width=40, height=2, bg="#ffffff", font=("Segoe UI", 10, "bold"), relief="ridge")
        self.btn_pga.pack(side="right", padx=10, expand=True)

        frame_acoes = tk.Frame(self, bg="black")
        frame_acoes.pack(fill="x", pady=15, padx=20)

        self.btn_processar = tk.Button(frame_acoes, text="Executar Comparação Automática", command=self.processar, bg="#2563eb", fg="white", font=("Segoe UI", 12, "bold"), height=2, relief="flat", cursor="hand2")
        self.btn_processar.pack(side="left", expand=True, fill="x", padx=(0, 10))

        self.btn_resetar = tk.Button(frame_acoes, text="🔄 Resetar", command=self.resetar, bg="#e2e8f0", fg="#1e293b", font=("Segoe UI", 12, "bold"), height=2, relief="flat", cursor="hand2", width=15)
        self.btn_resetar.pack(side="right")

        self.frame_resumo = tk.Frame(self, bg="black")

        cards_frame = tk.Frame(self.frame_resumo, bg="black")
        cards_frame.pack(side="left", fill="x", expand=True)

        self.lbl_card_prod = tk.Label(cards_frame, text="0\nProdutos", font=("Segoe UI", 11), bg="#f1f5f9", relief="groove")
        self.lbl_card_prod.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        self.lbl_card_ok = tk.Label(cards_frame, text="0\nBateram OK", font=("Segoe UI", 11, "bold"), fg="#166534", bg="#dcfce7", relief="groove")
        self.lbl_card_ok.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        self.lbl_card_err = tk.Label(cards_frame, text="0\nDivergências", font=("Segoe UI", 11, "bold"), fg="#991b1b", bg="#fee2e2", relief="groove")
        self.lbl_card_err.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        self.lbl_card_totpga = tk.Label(cards_frame, text="0,000\nTotal PGA (kg)", font=("Segoe UI", 11), bg="#f1f5f9", relief="groove")
        self.lbl_card_totpga.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        self.lbl_card_toterp = tk.Label(cards_frame, text="0,000\nTotal ERP (kg)", font=("Segoe UI", 11), bg="#f1f5f9", relief="groove")
        self.lbl_card_toterp.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        self.btn_export = tk.Button(self.frame_resumo, text="⬇ Exportar CSV", command=self.export_csv, bg="#ffffff", fg="#2563eb", font=("Segoe UI", 10, "bold"), relief="solid")
        self.btn_export.pack(side="right", padx=(10, 0), pady=5)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Producao.Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#f1f5f9", foreground="black")
        style.configure("Producao.Treeview", background="white", fieldbackground="white", foreground="black")

        self.tree = ttk.Treeview(self, columns=("produto", "pga", "erp", "diff", "status"), show="headings", height=12, style="Producao.Treeview")
        self.tree.heading("produto", text="Produto")
        self.tree.heading("pga", text="Total PGA (kg)")
        self.tree.heading("erp", text="Total ERP (kg)")
        self.tree.heading("diff", text="Diferença (kg)")
        self.tree.heading("status", text="Situação")

        self.tree.column("produto", width=350)
        self.tree.column("pga", width=110, anchor="e")
        self.tree.column("erp", width=110, anchor="e")
        self.tree.column("diff", width=110, anchor="e")
        self.tree.column("status", width=200, anchor="center")

        self.tree.pack(fill="both", expand=True, pady=10, padx=20)

    def carregar_erp(self):
        self.erp_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if self.erp_path:
            self.btn_erp.config(text=f"✅ ERP: {os.path.basename(self.erp_path)}", bg="#dcfce7")

    def carregar_pga(self):
        self.pga_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if self.pga_path:
            self.btn_pga.config(text=f"✅ PGA: {os.path.basename(self.pga_path)}", bg="#dcfce7")

    def resetar(self):
        self.erp_path = None
        self.pga_path = None
        self.last_results = []
        self.warnings_list = []

        self.btn_erp.config(text="📁 1. Selecionar ERP Ordenado", bg="#ffffff")
        self.btn_pga.config(text="📁 2. Selecionar PGA", bg="#ffffff")
        self.somente_divergencias.set(False)

        self.lbl_card_prod.config(text="0\nProdutos")
        self.lbl_card_ok.config(text="0\nOK")
        self.lbl_card_err.config(text="0\nDivergências")
        self.lbl_card_totpga.config(text="0,000\nTotal PGA (kg)")
        self.lbl_card_toterp.config(text="0,000\nTotal ERP (kg)")
        self.frame_resumo.pack_forget()

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.btn_processar.config(text="Executar Comparação Automática", state="normal")

    def atualizar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not self.last_results:
            return

        for r in self.last_results:
            if self.somente_divergencias.get() and r["sit"] == "OK":
                continue

            sit_texto = "✔ OK" if r['sit'] == 'OK' else (f"🚨 Falta no ERP" if r['sit'] == 'Falta no ERP' else (f"🚨 Falta no PGA" if r['sit'] == 'Falta no PGA' else f"🚨 Dif: {format_br(r['diff'])} kg"))
            valores = (r['produto'], format_br(r['pga']), format_br(r['erp']), format_br(r['diff']), sit_texto)
            self.tree.insert("", tk.END, values=valores)

    def processar(self):
        if not self.erp_path or not self.pga_path:
            messagebox.showwarning("Aviso", "Selecione as duas planilhas antes de continuar.")
            return

        self.btn_processar.config(text="Processando... aguarde", state="disabled")
        self.update_idletasks()
        self.warnings_list.clear()

        try:
            df_erp = pd.read_excel(self.erp_path, header=None, engine="openpyxl")
            df_pga = pd.read_excel(self.pga_path, header=None, engine="openpyxl")

            erp_r_prod, erp_c_prod = find_column(df_erp, [['descricao pga'], ['descricao'], ['produto']])
            erp_r_weight, erp_c_weight = find_column(df_erp, [['peso'], ['total']])

            linhas_header_erp = [r for r in (erp_r_prod, erp_r_weight) if r is not None]
            erp_header_row = max(linhas_header_erp) if linhas_header_erp else 1

            if erp_c_prod is None:
                erp_c_prod = 1
                self.warnings_list.append("ERP: Coluna 'Produto' não encontrada. Usando padrão (Coluna B).")
            if erp_c_weight is None:
                erp_c_weight = 3
                self.warnings_list.append("ERP: Coluna 'Peso' não encontrada. Usando padrão (Coluna D).")

            pga_r_prod, pga_c_prod = find_column(df_pga, [['produto']])
            pga_r_weight, pga_c_weight = find_column(df_pga, [['total'], ['peso']])

            linhas_header_pga = [r for r in (pga_r_prod, pga_r_weight) if r is not None]
            pga_header_row = max(linhas_header_pga) if linhas_header_pga else 0

            if pga_c_prod is None:
                pga_c_prod = 9
                self.warnings_list.append("PGA: Coluna 'Produto' não encontrada. Usando padrão (Coluna J).")
            if pga_c_weight is None:
                pga_c_weight = 12
                self.warnings_list.append("PGA: Coluna 'Total' não encontrada. Usando padrão (Coluna M).")

            erp_totals = {}
            for i in range(erp_header_row + 1, len(df_erp)):
                prod = normalize_product(df_erp.iloc[i, erp_c_prod])
                if prod and prod not in ['DESCRIÇÃO PGA', 'PRODUTO']:
                    val = limpar_numero_simples(df_erp.iloc[i, erp_c_weight])
                    erp_totals[prod] = erp_totals.get(prod, 0) + val

            pga_totals = {}
            for i in range(pga_header_row + 1, len(df_pga)):
                prod = normalize_product(df_pga.iloc[i, pga_c_prod])
                if prod and prod != 'PRODUTO':
                    val = limpar_numero_simples(df_pga.iloc[i, pga_c_weight])
                    pga_totals[prod] = pga_totals.get(prod, 0) + val

            all_products = set(list(erp_totals.keys()) + list(pga_totals.keys()))
            self.last_results = []
            count_ok, count_err = 0, 0
            total_pga_geral, total_erp_geral = 0.0, 0.0

            for p in sorted(list(all_products)):
                e = erp_totals.get(p, 0)
                pg = pga_totals.get(p, 0)
                diff = round(pg - e, 4)

                if p not in erp_totals:
                    sit = 'Falta no ERP'
                elif p not in pga_totals:
                    sit = 'Falta no PGA'
                elif abs(diff) <= TOLERANCIA_KG:
                    sit = 'OK'
                else:
                    sit = 'Divergência'

                if sit == 'OK':
                    count_ok += 1
                else:
                    count_err += 1
                total_pga_geral += pg
                total_erp_geral += e

                if sit == "Falta no ERP":
                    motivo = "Produto encontrado apenas no PGA"
                elif sit == "Falta no PGA":
                    motivo = "Produto encontrado apenas no ERP"
                elif sit == "Divergência":
                    motivo = f"Diferença de {format_br(diff)} kg"
                else:
                    motivo = "Pesos conciliados"

                self.last_results.append({
                    'produto': p, 'pga': pg, 'erp': e, 'diff': diff, 'sit': sit, 'motivo': motivo
                })

            self.frame_resumo.pack(fill="x", pady=5, padx=20)
            self.lbl_card_prod.config(text=f"{len(all_products)}\nProdutos")
            self.lbl_card_ok.config(text=f"{count_ok}\nOK")
            self.lbl_card_err.config(text=f"{count_err}\nDivergências")
            self.lbl_card_totpga.config(text=f"{format_br(total_pga_geral)}\nTotal PGA (kg)")
            self.lbl_card_toterp.config(text=f"{format_br(total_erp_geral)}\nTotal ERP (kg)")

            self.atualizar_tabela()

            if self.warnings_list:
                messagebox.showwarning("Aviso", "\n".join(self.warnings_list))

        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Erro crítico:\n{str(e)}")
        finally:
            self.btn_processar.config(text="Executar Comparação Automática", state="normal")

    def export_csv(self):
        if not self.last_results:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], title="Salvar Relatório")
        if not file_path:
            return

        with open(file_path, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['Produto', 'Total PGA (kg)', 'Total ERP (kg)', 'Diferenca (kg)', 'Situacao', 'Motivo'])
            for r in self.last_results:
                writer.writerow([r['produto'], f"{r['pga']:.3f}".replace('.', ','), f"{r['erp']:.3f}".replace('.', ','), f"{r['diff']:.3f}".replace('.', ','), r['sit'], r['motivo']])
        messagebox.showinfo("Sucesso", "Planilha salva com sucesso!")


# ==========================================================================
# MÓDULO: RECEBIMENTO
# ==========================================================================

class RecebimentoFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="black")
        self.controller = controller
        self.erp_path = None
        self.pga_path = None
        self.last_results = []
        self.warnings_list = []
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="SIFStat (Recebimento)", font=("Segoe UI", 18, "bold"), bg="black", fg="white").pack(pady=(20, 5))
        tk.Label(self, text="Cruza ERP e PGA agrupando por Produto e SIF de Origem.", font=("Segoe UI", 10), bg="black", fg="#cbd5e1").pack(pady=(0, 15))

        self.somente_divergencias = tk.BooleanVar(value=False)
        tk.Checkbutton(
            self, text="Mostrar apenas divergências", variable=self.somente_divergencias,
            bg="black", fg="white", selectcolor="black", activebackground="black", activeforeground="white",
            font=("Segoe UI", 10), command=self.atualizar_tabela
        ).pack(pady=(0, 10))

        frame_upload = tk.Frame(self, bg="black")
        frame_upload.pack(fill="x", pady=10, padx=20)

        self.btn_erp = tk.Button(frame_upload, text="📁 1. Selecionar ERP Recebimento", command=self.carregar_erp, width=40, height=2, bg="#ffffff", font=("Segoe UI", 10, "bold"), relief="ridge")
        self.btn_erp.pack(side="left", padx=10, expand=True)

        self.btn_pga = tk.Button(frame_upload, text="📁 2. Selecionar PGA Recebimento", command=self.carregar_pga, width=40, height=2, bg="#ffffff", font=("Segoe UI", 10, "bold"), relief="ridge")
        self.btn_pga.pack(side="right", padx=10, expand=True)

        frame_acoes = tk.Frame(self, bg="black")
        frame_acoes.pack(fill="x", pady=15, padx=20)

        self.btn_processar = tk.Button(frame_acoes, text="Executar Comparação Automática", command=self.processar, bg="#2563eb", fg="white", font=("Segoe UI", 12, "bold"), height=2, relief="flat", cursor="hand2")
        self.btn_processar.pack(side="left", expand=True, fill="x", padx=(0, 10))

        self.btn_resetar = tk.Button(frame_acoes, text="🔄 Resetar", command=self.resetar, bg="#e2e8f0", fg="#1e293b", font=("Segoe UI", 12, "bold"), height=2, relief="flat", cursor="hand2", width=15)
        self.btn_resetar.pack(side="right")

        self.frame_resumo = tk.Frame(self, bg="black")

        for i in range(6):
            self.frame_resumo.columnconfigure(i, weight=1)
        self.frame_resumo.columnconfigure(6, weight=0)

        self.lbl_card_prod = tk.Label(self.frame_resumo, text="0\nLinhas", font=("Segoe UI", 10), bg="#f1f5f9", relief="groove")
        self.lbl_card_prod.grid(row=0, column=0, sticky="ew", padx=3, pady=2)

        self.lbl_card_ok = tk.Label(self.frame_resumo, text="0\nBateram OK", font=("Segoe UI", 10, "bold"), fg="#166534", bg="#dcfce7", relief="groove")
        self.lbl_card_ok.grid(row=0, column=1, sticky="ew", padx=3, pady=2)

        self.lbl_card_err = tk.Label(self.frame_resumo, text="0\nDivergências", font=("Segoe UI", 10, "bold"), fg="#991b1b", bg="#fee2e2", relief="groove")
        self.lbl_card_err.grid(row=0, column=2, sticky="ew", padx=3, pady=2)

        self.lbl_card_ign = tk.Label(self.frame_resumo, text="0\nIgnorados", font=("Segoe UI", 10, "bold"), fg="#854d0e", bg="#fef08a", relief="groove")
        self.lbl_card_ign.grid(row=0, column=3, sticky="ew", padx=3, pady=2)

        self.lbl_card_totpga = tk.Label(self.frame_resumo, text="0,000\nTotal PGA (kg)", font=("Segoe UI", 10), bg="#f1f5f9", relief="groove")
        self.lbl_card_totpga.grid(row=0, column=4, sticky="ew", padx=3, pady=2)

        self.lbl_card_toterp = tk.Label(self.frame_resumo, text="0,000\nTotal ERP (kg)", font=("Segoe UI", 10), bg="#f1f5f9", relief="groove")
        self.lbl_card_toterp.grid(row=0, column=5, sticky="ew", padx=3, pady=2)

        self.btn_export = tk.Button(self.frame_resumo, text="⬇ Exportar CSV", command=self.export_csv, bg="#ffffff", fg="#2563eb", font=("Segoe UI", 10, "bold"), relief="solid")
        self.btn_export.grid(row=0, column=6, sticky="e", padx=(10, 0))

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Recebimento.Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#f1f5f9", foreground="black")
        style.configure("Recebimento.Treeview", background="white", fieldbackground="white", foreground="black")

        self.tree = ttk.Treeview(self, columns=("produto", "sif", "pga", "erp", "diff", "status"), show="headings", height=12, style="Recebimento.Treeview")
        self.tree.heading("produto", text="Produto")
        self.tree.heading("sif", text="SIF Origem")
        self.tree.heading("pga", text="Total PGA (kg)")
        self.tree.heading("erp", text="Total ERP (kg)")
        self.tree.heading("diff", text="Diferença (kg)")
        self.tree.heading("status", text="Situação")

        self.tree.column("produto", width=300)
        self.tree.column("sif", width=100, anchor="center")
        self.tree.column("pga", width=100, anchor="e")
        self.tree.column("erp", width=100, anchor="e")
        self.tree.column("diff", width=100, anchor="e")
        self.tree.column("status", width=200, anchor="center")

        self.tree.pack(fill="both", expand=True, pady=10, padx=20)

    def carregar_erp(self):
        self.erp_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if self.erp_path:
            self.btn_erp.config(text=f"✅ ERP: {os.path.basename(self.erp_path)}", bg="#dcfce7")

    def carregar_pga(self):
        self.pga_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if self.pga_path:
            self.btn_pga.config(text=f"✅ PGA: {os.path.basename(self.pga_path)}", bg="#dcfce7")

    def resetar(self):
        self.erp_path = None
        self.pga_path = None
        self.last_results = []
        self.warnings_list = []
        self.btn_erp.config(text="📁 1. Selecionar ERP Recebimento", bg="#ffffff")
        self.btn_pga.config(text="📁 2. Selecionar PGA Recebimento", bg="#ffffff")
        self.somente_divergencias.set(False)

        self.lbl_card_prod.config(text="0\nLinhas")
        self.lbl_card_ok.config(text="0\nOK")
        self.lbl_card_err.config(text="0\nDivergências")
        self.lbl_card_ign.config(text="0\nIgnorados")
        self.lbl_card_totpga.config(text="0,000\nTotal PGA (kg)")
        self.lbl_card_toterp.config(text="0,000\nTotal ERP (kg)")
        self.frame_resumo.pack_forget()

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.btn_processar.config(text="Executar Comparação Automática", state="normal")

    def atualizar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if not self.last_results:
            return
        for r in self.last_results:
            if self.somente_divergencias.get() and r["sit"] in ["OK", "Não necessário para conferência"]:
                continue

            sit_texto = "✔ OK" if r['sit'] == 'OK' else (
                "🟡 Ñ Necessário" if r['sit'] == 'Não necessário para conferência' else (
                    f"🚨 Falta no ERP" if r['sit'] == 'Falta no ERP' else (
                        f"🚨 Falta no PGA" if r['sit'] == 'Falta no PGA' else
                        f"🚨 Dif: {format_br(r['diff'])} kg")))

            valores = (r['produto'], r['sif'], format_br(r['pga']), format_br(r['erp']), format_br(r['diff']), sit_texto)
            self.tree.insert("", tk.END, values=valores)

    def processar(self):
        if not self.erp_path or not self.pga_path:
            messagebox.showwarning("Aviso", "Selecione as duas planilhas antes de continuar.")
            return

        self.btn_processar.config(text="Processando... aguarde", state="disabled")
        self.update_idletasks()
        self.warnings_list.clear()

        try:
            df_erp = pd.read_excel(self.erp_path, header=None, engine="openpyxl")
            df_pga = pd.read_excel(self.pga_path, header=None, engine="openpyxl")

            erp_r_prod, erp_c_prod = find_column(df_erp, [['descricao pga', 'descricao', 'produto']])
            erp_r_tipo, erp_c_tipo = find_column(df_erp, [['recebimento', 'tipo', 'operacao']])
            erp_r_sif, erp_c_sif = find_column(df_erp, [['sif origem', 'sif']])
            erp_r_weight, erp_c_weight = find_column(df_erp, [['peso', 'total']])

            linhas_header_erp = [r for r in (erp_r_prod, erp_r_tipo, erp_r_sif, erp_r_weight) if r is not None]
            erp_header_row = max(linhas_header_erp) if linhas_header_erp else 1

            if erp_c_prod is None:
                erp_c_prod = 2
                self.warnings_list.append("ERP: Coluna 'Produto' não encontrada. Usando padrão (Coluna C).")
            if erp_c_tipo is None:
                erp_c_tipo = 5
                self.warnings_list.append("ERP: Coluna 'Tipo' não encontrada. Usando padrão (Coluna F).")
            if erp_c_sif is None:
                erp_c_sif = 6
                self.warnings_list.append("ERP: Coluna 'SIF' não encontrada. Usando padrão (Coluna G).")
            if erp_c_weight is None:
                erp_c_weight = 7
                self.warnings_list.append("ERP: Coluna 'Peso' não encontrada. Usando padrão (Coluna H).")

            pga_r_prod, pga_c_prod = find_column(df_pga, [['produto', 'descricao']])
            pga_r_op, pga_c_op = find_column(df_pga, [['operação', 'operacao', 'tipo']])
            pga_r_sif, pga_c_sif = find_column(df_pga, [['controle origem (receb)', 'controle origem', 'sif origem', 'sif']])
            pga_r_weight, pga_c_weight = find_column(df_pga, [['total', 'peso']])

            linhas_header_pga = [r for r in (pga_r_prod, pga_r_op, pga_r_sif, pga_r_weight) if r is not None]
            pga_header_row = max(linhas_header_pga) if linhas_header_pga else 0

            if pga_c_op is None:
                pga_c_op = 0
                self.warnings_list.append("PGA: Coluna 'Operação' não encontrada. Usando padrão (Coluna A).")
            if pga_c_prod is None:
                pga_c_prod = 8
                self.warnings_list.append("PGA: Coluna 'Produto' não encontrada. Usando padrão (Coluna I).")
            if pga_c_sif is None:
                pga_c_sif = 10
                self.warnings_list.append("PGA: Coluna 'Controle Origem' não encontrada. Usando padrão (Coluna K).")
            if pga_c_weight is None:
                pga_c_weight = 14
                self.warnings_list.append("PGA: Coluna 'Total' não encontrada. Usando padrão (Coluna O).")

            erp_totals = {}
            erp_ignorados = {}
            for i in range(erp_header_row + 1, len(df_erp)):
                prod = normalize_product(df_erp.iloc[i, erp_c_prod])
                if prod and prod not in ['DESCRIÇÃO PGA', 'PRODUTO']:
                    sif = normalize_sif(df_erp.iloc[i, erp_c_sif])
                    val = round(limpar_numero_regex(df_erp.iloc[i, erp_c_weight]), 4)
                    chave = (prod, sif)

                    tipo_op = str(df_erp.iloc[i, erp_c_tipo]).upper()

                    if "ESTABELECIMENTO" in tipo_op:
                        erp_totals[chave] = erp_totals.get(chave, 0) + val
                    else:
                        erp_ignorados[chave] = erp_ignorados.get(chave, 0) + val

            pga_totals = {}
            pga_ignorados = {}
            for i in range(pga_header_row + 1, len(df_pga)):
                prod = normalize_product(df_pga.iloc[i, pga_c_prod])
                if prod and prod != 'PRODUTO':
                    sif = normalize_sif(df_pga.iloc[i, pga_c_sif])
                    val = round(limpar_numero_regex(df_pga.iloc[i, pga_c_weight]), 4)
                    chave = (prod, sif)

                    op_pga = str(df_pga.iloc[i, pga_c_op]).upper()

                    if "AUTORIZADO" in op_pga:
                        pga_ignorados[chave] = pga_ignorados.get(chave, 0) + val
                    else:
                        pga_totals[chave] = pga_totals.get(chave, 0) + val

            self.last_results = []
            count_ok, count_err, count_ign = 0, 0, 0
            total_pga_geral, total_erp_geral = 0.0, 0.0

            todas_chaves_validas = set(list(erp_totals.keys()) + list(pga_totals.keys()))
            for chave in sorted(list(todas_chaves_validas)):
                p, sif = chave
                e = erp_totals.get(chave, 0)
                pg = pga_totals.get(chave, 0)

                diff = round(pg - e, 4)

                if chave not in erp_totals:
                    sit = 'Falta no ERP'
                elif chave not in pga_totals:
                    sit = 'Falta no PGA'
                elif abs(diff) <= TOLERANCIA_KG:
                    sit = 'OK'
                else:
                    sit = 'Divergência'

                if sit == 'OK':
                    count_ok += 1
                else:
                    count_err += 1

                total_pga_geral += pg
                total_erp_geral += e

                if sit == "Falta no ERP":
                    motivo = "Recebimento do SIF consta apenas no PGA"
                elif sit == "Falta no PGA":
                    motivo = "Recebimento do SIF consta apenas no ERP"
                elif sit == "Divergência":
                    motivo = f"Diferença de {format_br(diff)} kg"
                else:
                    motivo = "Pesos conciliados"

                self.last_results.append({
                    'produto': p, 'sif': sif, 'pga': pg, 'erp': e, 'diff': diff, 'sit': sit, 'motivo': motivo
                })

            todas_chaves_ignoradas = set(list(erp_ignorados.keys()) + list(pga_ignorados.keys()))
            for chave in sorted(list(todas_chaves_ignoradas)):
                p, sif = chave
                e = erp_ignorados.get(chave, 0)
                pg = pga_ignorados.get(chave, 0)
                diff = round(pg - e, 4)

                self.last_results.append({
                    'produto': p, 'sif': sif, 'pga': pg, 'erp': e, 'diff': diff,
                    'sit': 'Não necessário para conferência',
                    'motivo': 'Receb. Autorizado ou Devolução'
                })
                count_ign += 1

            self.frame_resumo.pack(fill="x", pady=5, padx=20)
            self.lbl_card_prod.config(text=f"{len(todas_chaves_validas) + count_ign}\nLinhas")
            self.lbl_card_ok.config(text=f"{count_ok}\nOK")
            self.lbl_card_err.config(text=f"{count_err}\nDivergências")
            self.lbl_card_ign.config(text=f"{count_ign}\nIgnorados")
            self.lbl_card_totpga.config(text=f"{format_br(total_pga_geral)}\nTotal PGA (kg)")
            self.lbl_card_toterp.config(text=f"{format_br(total_erp_geral)}\nTotal ERP (kg)")

            self.atualizar_tabela()

            if self.warnings_list:
                messagebox.showwarning("Aviso", "\n".join(self.warnings_list))

        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Erro crítico:\n{str(e)}")
        finally:
            self.btn_processar.config(text="Executar Comparação Automática", state="normal")

    def export_csv(self):
        if not self.last_results:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], title="Salvar Relatório")
        if not file_path:
            return

        with open(file_path, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['Produto', 'SIF Origem', 'Total PGA (kg)', 'Total ERP (kg)', 'Diferenca (kg)', 'Situacao', 'Motivo'])
            for r in self.last_results:
                writer.writerow([r['produto'], r['sif'], f"{r['pga']:.3f}".replace('.', ','), f"{r['erp']:.3f}".replace('.', ','), f"{r['diff']:.3f}".replace('.', ','), r['sit'], r['motivo']])
        messagebox.showinfo("Sucesso", "Planilha salva com sucesso!")


# ==========================================================================
# MÓDULO: COMERCIALIZAÇÃO
# ==========================================================================

DE_PARA_PAISES_BASE = {
    'SEBIA AND COSOVO': 'SERVIA',
    'ABW': 'ARUBA', 'AFG': 'AFEGANISTAO', 'AGO': 'ANGOLA', 'AIA': 'ANGUILLA', 'ALA': 'ILHAS ALAND',
    'ALB': 'ALBANIA', 'AND': 'ANDORRA', 'ARE': 'EMIRADOS ARABES UNIDOS', 'ARG': 'ARGENTINA',
    'ARM': 'ARMENIA', 'ASM': 'SAMOA AMERICANA', 'ATA': 'ANTARTIDA', 'ATF': 'TERRITORIOS DO SUL DA FRANCA',
    'ATG': 'ANTIGUA E BARBUDA', 'AUS': 'AUSTRALIA', 'AUT': 'AUSTRIA', 'AZE': 'AZERBAIJAO', 'BDI': 'BURUNDI',
    'BEL': 'BELGICA', 'BEN': 'BENIN', 'BES': 'SAO BARTOLOMEU', 'BFA': 'BURKINA FASO', 'BGD': 'BANGLADESH',
    'BGR': 'BULGARIA', 'BHR': 'BAHREIN', 'BHS': 'BAHAMAS', 'BIH': 'BOSNIA E HERZEGOVINA', 'BLM': 'SAO BARTOLOMEU',
    'BLR': 'BIELORRUSSIA', 'BLZ': 'BELIZE', 'BMU': 'BERMUDAS', 'BOL': 'BOLIVIA', 'BRA': 'BRASIL',
    'BRB': 'BARBADOS', 'BRN': 'BRUNEI', 'BTN': 'BUTAO', 'BVT': 'ILHA BOUVET', 'BWA': 'BOTSUANA',
    'CAF': 'REPUBLICA CENTRO-AFRICANA', 'CAN': 'CANADA', 'CCK': 'ILHAS COCOS (KEELING)', 'CHE': 'SUICA',
    'CHL': 'CHILE', 'CHN': 'CHINA', 'CIV': 'COSTA DO MARFIM', 'CMR': 'CAMAROES', 'COD': 'REPUBLICA DEMOCRATICA DO CONGO',
    'COG': 'CONGO', 'COK': 'ILHAS COOK', 'COL': 'COLOMBIA', 'COM': 'COMORES', 'CPV': 'CABO VERDE',
    'CRI': 'COSTA RICA', 'CUB': 'CUBA', 'CUW': 'CURACAO', 'CXR': 'ILHA CHRISTMAS', 'CYM': 'ILHAS CAYMAN',
    'CYP': 'CHIPRE', 'CZE': 'REPUBLICA TCHECA', 'DEU': 'ALEMANHA', 'DJI': 'DJIBUTI', 'DMA': 'DOMINICA',
    'DNK': 'DINAMARCA', 'DOM': 'REPUBLICA DOMINICANA', 'DZA': 'ARGELIA', 'ECU': 'EQUADOR', 'EGY': 'EGITO',
    'ERI': 'ERITREIA', 'ESH': 'SAARA OCIDENTAL', 'ESP': 'ESPANHA', 'EST': 'ESTONIA', 'ETH': 'ETIOPIA',
    'FIN': 'FINLANDIA', 'FJI': 'FIJI', 'FLK': 'ILHAS FALKLAND', 'FRA': 'FRANCA', 'FRO': 'ILHAS FAROE',
    'FSM': 'MICRONESIA', 'GAB': 'GABAO', 'GBR': 'REINO UNIDO', 'GEO': 'GEORGIA', 'GGY': 'GUERNSEY',
    'GHA': 'GANA', 'GIB': 'GIBRALTAR', 'GIN': 'GUINE', 'GLP': 'GUADALUPE', 'GMB': 'GAMBIA',
    'GNB': 'GUINE-BISSAU', 'GNQ': 'GUINE EQUATORIAL', 'GRC': 'GRECIA', 'GRD': 'GRANADA', 'GRL': 'GROENLANDIA',
    'GTM': 'GUATEMALA', 'GUF': 'GUIANA FRANCESA', 'GUM': 'GUAM', 'GUY': 'GUIANA', 'HKG': 'HONG KONG',
    'HMD': 'ILHAS HEARD E MCDONALD', 'HND': 'HONDURAS', 'HRV': 'CROACIA', 'HTI': 'HAITI', 'HUN': 'HUNGRIA',
    'IDN': 'INDONESIA', 'IMN': 'ILHA DE MAN', 'IND': 'INDIA', 'IOT': 'TERRITORIO BRITANICO DO OCEANO INDICO',
    'IRL': 'IRLANDA', 'IRN': 'IRA', 'IRQ': 'IRAQUE', 'ISL': 'ISLANDIA', 'ISR': 'ISRAEL', 'ITA': 'ITALIA',
    'JAM': 'JAMAICA', 'JEY': 'JERSEY', 'JOR': 'JORDANIA', 'JPN': 'JAPAO', 'KAZ': 'CAZAQUISTAO',
    'KEN': 'QUENIA', 'KGZ': 'QUIRGUISTAO', 'KHM': 'CAMBOJA', 'KIR': 'KIRIBATI', 'KNA': 'SAO CRISTOVAO E NEVIS',
    'KOR': 'COREIA DO SUL', 'KWT': 'KUWAIT', 'LAO': 'LAOS', 'LBN': 'LIBANO', 'LBR': 'LIBERIA',
    'LBY': 'LIBIA', 'LCA': 'SANTA LUCIA', 'LIE': 'LIECHTENSTEIN', 'LKA': 'SRI LANKA', 'LSO': 'LESOTO',
    'LTU': 'LITUANIA', 'LUX': 'LUXEMBURGO', 'LVA': 'LETONIA', 'MAC': 'MACAU', 'MAF': 'SAO MARTINHO',
    'MAR': 'MARROCOS', 'MCO': 'MONACO', 'MDA': 'MOLDAVIA', 'MDG': 'MADAGASCAR', 'MDV': 'MALDIVAS',
    'MEX': 'MEXICO', 'MHL': 'ILHAS MARSHALL', 'MKD': 'MACEDONIA DO NORTE', 'MLI': 'MALI', 'MLT': 'MALTA',
    'MMR': 'MIANMAR', 'MNE': 'MONTENEGRO', 'MNG': 'MONGOLIA', 'MNP': 'ILHAS MARIANAS DO NORTE',
    'MOZ': 'MOCAMBIQUE', 'MRT': 'MAURITANIA', 'MSR': 'MONTSERRAT', 'MTQ': 'MARTINICA', 'MUS': 'MAURICIO',
    'MWI': 'MALAWI', 'MYS': 'MALASIA', 'MYT': 'MAYOTTE', 'NAM': 'NAMIBIA', 'NCL': 'NOVA CALEDONIA',
    'NER': 'NIGER', 'NFK': 'ILHA NORFOLK', 'NGA': 'NIGERIA', 'NIC': 'NICARAGUA', 'NIU': 'NIUE',
    'NLD': 'HOLANDA (PAISES BAIXOS)', 'NOR': 'NORUEGA', 'NPL': 'NEPAL', 'NRU': 'NAURU', 'NZL': 'NOVA ZELANDIA',
    'OMN': 'OMA', 'PAK': 'PAQUISTAO', 'PAN': 'PANAMA', 'PCN': 'ILHAS PITCAIRN', 'PER': 'PERU',
    'PHL': 'FILIPINAS', 'PLW': 'PALAU', 'PNG': 'PAPUA NOVA GUINE', 'POL': 'POLONIA', 'PRI': 'PORTO RICO',
    'PRK': 'COREIA DO NORTE', 'PRT': 'PORTUGAL', 'PRY': 'PARAGUAI', 'PSE': 'PALESTINA', 'PYF': 'POLINESIA FRANCESA',
    'QAT': 'CATAR', 'REU': 'REUNIAO', 'ROU': 'ROMENIA', 'RUS': 'RUSSIA', 'RWA': 'RUANDA', 'SAU': 'ARABIA SAUDITA',
    'SDN': 'SUDAO', 'SEN': 'SENEGAL', 'SGP': 'CINGAPURA', 'SGS': 'ILHAS GEORGIA DO SUL E SANDWICH DO SUL',
    'SHN': 'SANTA HELENA', 'SJM': 'SVALBARD E JAN MAYEN', 'SLB': 'ILHAS SALOMAO', 'SLE': 'SERRA LEOA',
    'SLV': 'EL SALVADOR', 'SMR': 'SAN MARINO', 'SOM': 'SOMALIA', 'SPM': 'SAO PEDRO E MIQUELON',
    'SRB': 'SERVIA', 'SSD': 'SUDAO DO SUL', 'STP': 'SAO TOME E PRINCIPE', 'SUR': 'SURINAME', 'SVK': 'ESLOVAQUIA',
    'SVN': 'ESLOVENIA', 'SWE': 'SUECIA', 'SWZ': 'SUAZILANDIA', 'SXM': 'SINT MAARTEN', 'SYC': 'SEICHELES',
    'SYR': 'SIRIA', 'TCA': 'ILHAS TURCAS E CAICOS', 'TCD': 'CHADE', 'TGO': 'TOGO', 'THA': 'TAILANDIA',
    'TJK': 'TAJIQUISTAO', 'TKL': 'TOKELAU', 'TKM': 'TURCOMENISTAO', 'TLS': 'TIMOR-LESTE', 'TON': 'TONGA',
    'TTO': 'TRINIDAD E TOBAGO', 'TUN': 'TUNISIA', 'TUR': 'TURQUIA', 'TUV': 'TUVALU', 'TWN': 'TAIWAN',
    'TZA': 'TANZANIA', 'UGA': 'UGANDA', 'UKR': 'UCRANIA', 'UMI': 'ILHAS MENORES DISTANTES DOS EUA',
    'URY': 'URUGUAI', 'USA': 'ESTADOS UNIDOS', 'UZB': 'UZBEQUISTAO', 'VAT': 'VATICANO',
    'VCT': 'SAO VICENTE E GRANADINAS', 'VEN': 'VENEZUELA', 'VGB': 'ILHAS VIRGENS BRITANICAS',
    'VIR': 'ILHAS VIRGENS, EUA', 'VNM': 'VIETNA', 'VUT': 'VANUATU', 'WLF': 'WALLIS E FUTUNA',
    'WSM': 'SAMOA', 'YEM': 'IEMEN', 'ZAF': 'AFRICA DO SUL', 'ZMB': 'ZAMBIA', 'ZWE': 'ZIMBABUE'
}


def obter_caminho_config_comercializacao():
    """Garante que a planilha de memória de países seja salva ao lado do executável."""
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, 'Config_Paises_SIFStat.xlsx')


class ComercializacaoFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="black")
        self.controller = controller
        self.erp_path = None
        self.pga_path = None
        self.last_results = []
        self.warnings_list = []
        self.dicionario_customizado = self.carregar_dicionario_customizado()
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="SIFStat (Comercialização)", font=("Segoe UI", 18, "bold"), bg="black", fg="white").pack(pady=(20, 5))
        tk.Label(self, text="Faça o upload do ERP ordenado e do PGA para conciliar Pesos por Produto e UF/País.", font=("Segoe UI", 10), bg="black", fg="#cbd5e1").pack(pady=(0, 15))

        self.somente_divergencias = tk.BooleanVar(value=False)
        tk.Checkbutton(
            self, text="Mostrar apenas divergências", variable=self.somente_divergencias,
            bg="black", fg="white", selectcolor="black", activebackground="black", activeforeground="white",
            font=("Segoe UI", 10), command=self.atualizar_tabela
        ).pack(pady=(0, 10))

        frame_upload = tk.Frame(self, bg="black")
        frame_upload.pack(fill="x", pady=(10, 5), padx=20)

        self.btn_erp = tk.Button(frame_upload, text="📁 1. Selecionar ERP Comercialização", command=self.carregar_erp, width=40, height=2, bg="#ffffff", font=("Segoe UI", 10, "bold"), relief="ridge")
        self.btn_erp.pack(side="left", padx=10, expand=True)

        self.btn_pga = tk.Button(frame_upload, text="📁 2. Selecionar PGA Comercialização", command=self.carregar_pga, width=40, height=2, bg="#ffffff", font=("Segoe UI", 10, "bold"), relief="ridge")
        self.btn_pga.pack(side="right", padx=10, expand=True)

        frame_cadastro = tk.Frame(self, bg="black")
        frame_cadastro.pack(fill="x", pady=(5, 10), padx=20)

        self.btn_cadastro_pais = tk.Button(frame_cadastro, text="➕ Cadastrar / Editar País", command=self.abrir_janela_cadastro, height=1, bg="#fef08a", fg="#854d0e", font=("Segoe UI", 9, "bold"), relief="ridge")
        self.btn_cadastro_pais.pack(side="left", padx=(10, 5), expand=True, fill="x")

        self.btn_abrir_memoria = tk.Button(frame_cadastro, text="📂 Ver Planilha de Memória", command=self.abrir_planilha_memoria, height=1, bg="#e2e8f0", fg="#1e293b", font=("Segoe UI", 9, "bold"), relief="ridge")
        self.btn_abrir_memoria.pack(side="right", padx=(5, 10), expand=True, fill="x")

        frame_acoes = tk.Frame(self, bg="black")
        frame_acoes.pack(fill="x", pady=15, padx=20)

        self.btn_processar = tk.Button(frame_acoes, text="Executar Comparação Automática", command=self.processar, bg="#2563eb", fg="white", font=("Segoe UI", 12, "bold"), height=2, relief="flat", cursor="hand2")
        self.btn_processar.pack(side="left", expand=True, fill="x", padx=(0, 10))

        self.btn_resetar = tk.Button(frame_acoes, text="🔄 Resetar", command=self.resetar, bg="#e2e8f0", fg="#1e293b", font=("Segoe UI", 12, "bold"), height=2, relief="flat", cursor="hand2", width=15)
        self.btn_resetar.pack(side="right")

        self.frame_resumo = tk.Frame(self, bg="black")
        self.lbl_card_prod = tk.Label(self.frame_resumo, text="0\nLinhas", font=("Segoe UI", 11), bg="#f1f5f9", width=15, relief="groove")
        self.lbl_card_prod.pack(side="left", padx=5)
        self.lbl_card_ok = tk.Label(self.frame_resumo, text="0\nBateram OK", font=("Segoe UI", 11, "bold"), fg="#166534", bg="#dcfce7", width=15, relief="groove")
        self.lbl_card_ok.pack(side="left", padx=5)
        self.lbl_card_err = tk.Label(self.frame_resumo, text="0\nDivergências", font=("Segoe UI", 11, "bold"), fg="#991b1b", bg="#fee2e2", width=15, relief="groove")
        self.lbl_card_err.pack(side="left", padx=5)
        self.lbl_card_totpga = tk.Label(self.frame_resumo, text="0,000\nTotal PGA (kg)", font=("Segoe UI", 11), bg="#f1f5f9", width=18, relief="groove")
        self.lbl_card_totpga.pack(side="left", padx=5)
        self.lbl_card_toterp = tk.Label(self.frame_resumo, text="0,000\nTotal ERP (kg)", font=("Segoe UI", 11), bg="#f1f5f9", width=18, relief="groove")
        self.lbl_card_toterp.pack(side="left", padx=5)

        self.btn_export = tk.Button(self.frame_resumo, text="⬇ Exportar CSV", command=self.export_csv, bg="#ffffff", fg="#2563eb", font=("Segoe UI", 10, "bold"), relief="solid")
        self.btn_export.pack(side="right", padx=10)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Comercializacao.Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#f1f5f9", foreground="black")
        style.configure("Comercializacao.Treeview", background="white", fieldbackground="white", foreground="black")

        self.tree = ttk.Treeview(self, columns=("produto", "uf", "pga", "erp", "diff", "status"), show="headings", height=10, style="Comercializacao.Treeview")
        self.tree.heading("produto", text="Produto")
        self.tree.heading("uf", text="UF / País")
        self.tree.heading("pga", text="Total PGA (kg)")
        self.tree.heading("erp", text="Total ERP (kg)")
        self.tree.heading("diff", text="Diferença (kg)")
        self.tree.heading("status", text="Situação")
        self.tree.column("produto", width=300)
        self.tree.column("uf", width=100, anchor="center")
        self.tree.column("pga", width=100, anchor="e")
        self.tree.column("erp", width=100, anchor="e")
        self.tree.column("diff", width=100, anchor="e")
        self.tree.column("status", width=200, anchor="center")
        self.tree.pack(fill="both", expand=True, pady=10, padx=20)

    # ----------------------------------------------------------------
    # Cadastro / memória de países
    # ----------------------------------------------------------------
    def abrir_planilha_memoria(self):
        caminho = obter_caminho_config_comercializacao()
        try:
            if not os.path.exists(caminho):
                pd.DataFrame(columns=['NOME NO ERP', 'NOME NO PGA']).to_excel(caminho, index=False)
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo.\nCaminho: {caminho}\n\nErro: {e}")

    def abrir_janela_cadastro(self):
        janela = tk.Toplevel(self)
        janela.title("Cadastrar Novo País")
        janela.geometry("450x250")
        janela.configure(bg="#1e293b", padx=20, pady=20)
        janela.attributes("-topmost", True)
        janela.focus_force()

        try:
            janela.iconbitmap(resource_path(os.path.join("assets", "icone.ico")))
        except Exception:
            pass

        tk.Label(janela, text="Nome no ERP:", bg="#1e293b", fg="white", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        entry_erp = tk.Entry(janela, font=("Segoe UI", 12), width=40)
        entry_erp.pack(fill="x", pady=(0, 15))

        tk.Label(janela, text="Nome no PGA:", bg="#1e293b", fg="white", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        entry_pga = tk.Entry(janela, font=("Segoe UI", 12), width=40)
        entry_pga.pack(fill="x", pady=(0, 20))

        def salvar_novo_pais():
            erp_val = entry_erp.get().strip().upper()
            pga_val = entry_pga.get().strip().upper()

            if not erp_val or not pga_val:
                messagebox.showwarning("Atenção", "Preencha os dois campos antes de salvar!", parent=janela)
                return

            caminho_config = obter_caminho_config_comercializacao()
            salvo_excel = False

            try:
                if os.path.exists(caminho_config):
                    df_config = pd.read_excel(caminho_config)
                else:
                    df_config = pd.DataFrame(columns=['NOME NO ERP', 'NOME NO PGA'])

                df_config = df_config[df_config['NOME NO ERP'] != erp_val]
                nova_linha = pd.DataFrame([{'NOME NO ERP': erp_val, 'NOME NO PGA': pga_val}])
                df_config = pd.concat([df_config, nova_linha], ignore_index=True)
                df_config.to_excel(caminho_config, index=False)
                salvo_excel = True
            except Exception as e:
                messagebox.showwarning("Aviso", f"O país foi salvo na memória ativa do robô, mas não no arquivo Excel (ele pode estar aberto ou bloqueado).\nVocê já pode fazer o cruzamento sem problemas.\n\nDetalhe técnico: {e}", parent=janela)

            self.dicionario_customizado[remove_acentos(erp_val)] = remove_acentos(pga_val)

            if salvo_excel:
                messagebox.showinfo("Sucesso!", f"Tradução cadastrada com sucesso:\n\n{erp_val} ➜ {pga_val}\n\nO robô já decorou. Clique em Executar Comparação Automática!", parent=janela)
            janela.destroy()

        tk.Button(janela, text="💾 Salvar no Sistema", command=salvar_novo_pais, bg="#22c55e", fg="white", font=("Segoe UI", 11, "bold"), height=2, cursor="hand2").pack(fill="x")

    def carregar_dicionario_customizado(self):
        dic_custom = {}
        caminho_config = obter_caminho_config_comercializacao()
        if os.path.exists(caminho_config):
            try:
                df = pd.read_excel(caminho_config)
                for i in range(len(df)):
                    val_erp = str(df.iloc[i, 0]).strip().upper()
                    val_pga = str(df.iloc[i, 1]).strip().upper()
                    if val_erp and val_erp != 'NAN' and val_pga and val_pga != 'NAN':
                        dic_custom[remove_acentos(val_erp)] = remove_acentos(val_pga)
            except Exception:
                pass
        return dic_custom

    # ----------------------------------------------------------------
    # Funções básicas
    # ----------------------------------------------------------------
    def carregar_erp(self):
        self.erp_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if self.erp_path:
            self.btn_erp.config(text=f"✅ ERP: {os.path.basename(self.erp_path)}", bg="#dcfce7")

    def carregar_pga(self):
        self.pga_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if self.pga_path:
            self.btn_pga.config(text=f"✅ PGA: {os.path.basename(self.pga_path)}", bg="#dcfce7")

    def resetar(self):
        self.erp_path = None
        self.pga_path = None
        self.last_results = []
        self.warnings_list = []

        self.btn_erp.config(text="📁 1. Selecionar ERP Comercialização", bg="#ffffff")
        self.btn_pga.config(text="📁 2. Selecionar PGA Comercialização", bg="#ffffff")
        self.somente_divergencias.set(False)
        self.lbl_card_prod.config(text="0\nLinhas")
        self.lbl_card_ok.config(text="0\nOK")
        self.lbl_card_err.config(text="0\nDivergências")
        self.lbl_card_totpga.config(text="0,000\nTotal PGA (kg)")
        self.lbl_card_toterp.config(text="0,000\nTotal ERP (kg)")
        self.frame_resumo.pack_forget()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.btn_processar.config(text="Executar Comparação Automática", state="normal")

    def atualizar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if not self.last_results:
            return
        for r in self.last_results:
            if self.somente_divergencias.get() and r["sit"] == "OK":
                continue
            sit_texto = "✔ OK" if r['sit'] == 'OK' else (f"🚨 Falta no ERP" if r['sit'] == 'Falta no ERP' else (f"🚨 Falta no PGA" if r['sit'] == 'Falta no PGA' else f"🚨 Dif: {format_br(r['diff'])} kg"))
            valores = (r['produto'], r['uf'], format_br(r['pga']), format_br(r['erp']), format_br(r['diff']), sit_texto)
            self.tree.insert("", tk.END, values=valores)

    def processar(self):
        if not self.erp_path or not self.pga_path:
            messagebox.showwarning("Aviso", "Selecione as planilhas 1 e 2 antes de continuar.")
            return

        self.btn_processar.config(text="Processando... aguarde", state="disabled")
        self.update_idletasks()
        self.warnings_list.clear()

        try:
            dicionario_ativo = DE_PARA_PAISES_BASE.copy()
            dicionario_ativo.update(self.dicionario_customizado)

            df_erp = pd.read_excel(self.erp_path, header=None, engine="openpyxl")
            df_pga = pd.read_excel(self.pga_path, header=None, engine="openpyxl")

            erp_r_prod, erp_c_prod = find_column(df_erp, [['descricao pga'], ['descricao'], ['produto']])
            erp_r_uf, erp_c_uf = find_column(df_erp, [['destino'], ['uf']])
            erp_r_weight, erp_c_weight = find_column(df_erp, [['peso'], ['total']])

            linhas_header_erp = [r for r in (erp_r_prod, erp_r_uf, erp_r_weight) if r is not None]
            erp_header_row = max(linhas_header_erp) if linhas_header_erp else 1

            if erp_c_prod is None:
                erp_c_prod = 2
                self.warnings_list.append("ERP: Coluna 'Produto' não encontrada. Usando padrão (Coluna C).")
            if erp_c_uf is None:
                erp_c_uf = 4
                self.warnings_list.append("ERP: Coluna 'Destino' não encontrada. Usando padrão (Coluna E).")
            if erp_c_weight is None:
                erp_c_weight = 6
                self.warnings_list.append("ERP: Coluna 'Peso' não encontrada. Usando padrão (Coluna G).")

            pga_r_prod, pga_c_prod = find_column(df_pga, [['produto']])
            pga_r_uf, pga_c_uf = find_column(df_pga, [['uf/pais'], ['uf/país']])
            pga_r_weight, pga_c_weight = find_column(df_pga, [['total'], ['peso']])

            linhas_header_pga = [r for r in (pga_r_prod, pga_r_uf, pga_r_weight) if r is not None]
            pga_header_row = max(linhas_header_pga) if linhas_header_pga else 0

            if pga_c_prod is None:
                pga_c_prod = 8
                self.warnings_list.append("PGA: Coluna 'Produto' não encontrada. Usando padrão (Coluna I).")
            if pga_c_uf is None:
                pga_c_uf = 12
                self.warnings_list.append("PGA: Coluna 'UF/País' não encontrada. Usando padrão (Coluna M).")
            if pga_c_weight is None:
                pga_c_weight = 14
                self.warnings_list.append("PGA: Coluna 'Total' não encontrada. Usando padrão (Coluna O).")

            erp_totals = {}
            for i in range(erp_header_row + 1, len(df_erp)):
                prod = normalize_product_acentos(df_erp.iloc[i, erp_c_prod])
                uf = normalize_location(df_erp.iloc[i, erp_c_uf], dicionario_ativo)
                if prod and prod not in ['DESCRICAO PGA', 'PRODUTO']:
                    val = round(limpar_numero_regex(df_erp.iloc[i, erp_c_weight]), 4)
                    chave = (prod, uf)
                    erp_totals[chave] = erp_totals.get(chave, 0) + val

            pga_totals = {}
            for i in range(pga_header_row + 1, len(df_pga)):
                prod = normalize_product_acentos(df_pga.iloc[i, pga_c_prod])
                uf = normalize_location(df_pga.iloc[i, pga_c_uf], dicionario_ativo)
                if prod and prod != 'PRODUTO':
                    val = round(limpar_numero_regex(df_pga.iloc[i, pga_c_weight]), 4)
                    chave = (prod, uf)
                    pga_totals[chave] = pga_totals.get(chave, 0) + val

            todas_chaves = set(list(erp_totals.keys()) + list(pga_totals.keys()))
            self.last_results = []
            count_ok, count_err = 0, 0
            total_pga_geral, total_erp_geral = 0.0, 0.0

            for chave in sorted(list(todas_chaves)):
                p, uf = chave
                e = erp_totals.get(chave, 0)
                pg = pga_totals.get(chave, 0)
                diff = round(pg - e, 4)

                if chave not in erp_totals:
                    sit = 'Falta no ERP'
                elif chave not in pga_totals:
                    sit = 'Falta no PGA'
                elif abs(diff) <= TOLERANCIA_KG:
                    sit = 'OK'
                else:
                    sit = 'Divergência'

                if sit == 'OK':
                    count_ok += 1
                else:
                    count_err += 1
                total_pga_geral += pg
                total_erp_geral += e

                if sit == "Falta no ERP":
                    motivo = "Produto/UF encontrado apenas no PGA"
                elif sit == "Falta no PGA":
                    motivo = "Produto/UF encontrado apenas no ERP"
                elif sit == "Divergência":
                    motivo = f"Diferença de {format_br(diff)} kg"
                else:
                    motivo = "Pesos conciliados"

                self.last_results.append({
                    'produto': p, 'uf': uf, 'pga': pg, 'erp': e, 'diff': diff, 'sit': sit, 'motivo': motivo
                })

            self.frame_resumo.pack(fill="x", pady=5, padx=20)
            self.lbl_card_prod.config(text=f"{len(todas_chaves)}\nLinhas Cruzadas")
            self.lbl_card_ok.config(text=f"{count_ok}\nOK")
            self.lbl_card_err.config(text=f"{count_err}\nDivergências")
            self.lbl_card_totpga.config(text=f"{format_br(total_pga_geral)}\nTotal PGA (kg)")
            self.lbl_card_toterp.config(text=f"{format_br(total_erp_geral)}\nTotal ERP (kg)")

            self.atualizar_tabela()

            if self.warnings_list:
                messagebox.showwarning("Aviso", "\n".join(self.warnings_list))

        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Erro crítico:\n{str(e)}")
        finally:
            self.btn_processar.config(text="Executar Comparação Automática", state="normal")

    def export_csv(self):
        if not self.last_results:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], title="Salvar Relatório")
        if not file_path:
            return

        with open(file_path, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['Produto', 'UF / Pais', 'Total PGA (kg)', 'Total ERP (kg)', 'Diferenca (kg)', 'Situacao', 'Motivo'])
            for r in self.last_results:
                writer.writerow([r['produto'], r['uf'], f"{r['pga']:.3f}".replace('.', ','), f"{r['erp']:.3f}".replace('.', ','), f"{r['diff']:.3f}".replace('.', ','), r['sit'], r['motivo']])
        messagebox.showinfo("Sucesso", "Planilha salva com sucesso!")


# ==========================================================================
# PÁGINAS: INÍCIO / COMO UTILIZAR / CONTATO
# ==========================================================================

class HomePage(tk.Frame):
    """Página exibida ao abrir o programa."""
    def __init__(self, parent, controller):
        super().__init__(parent, bg="white")
        self.controller = controller

        wrapper = tk.Frame(self, bg="white")
        wrapper.place(relx=0.5, rely=0.42, anchor="center")

        tk.Label(wrapper, text="SIFSTAT", font=("Segoe UI", 28, "bold"), bg="white", fg="#111111").pack()
        tk.Label(
            wrapper,
            text="Selecione \"Mapas Estatísticos\" no menu ao lado e escolha\nProdução, Recebimento ou Comercialização para começar.",
            font=("Segoe UI", 11), bg="white", fg="#64748b", justify="center"
        ).pack(pady=(10, 0))


class ScrollableFrame(tk.Frame):
    """Frame com barra de rolagem vertical - usado para conteúdo longo (tutorial)."""

    def __init__(self, parent, bg="white"):
        super().__init__(parent, bg=bg)

        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=bg)

        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self._window, width=event.width)

    def _bind_mousewheel(self, _event=None):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event=None):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class ComoUtilizarPage(tk.Frame):
    """Página 'Como utilizar ?' - tutorial passo a passo (exportar ERP/PGA e comparar no SIFSTAT)."""

    CONTENT_WIDTH = 860
    TEXT_COLOR = "#1e293b"
    MUTED_COLOR = "#475569"
    ACCENT_COLOR = "#2563eb"
    WARNING_BG = "#fef9c3"
    WARNING_BORDER = "#eab308"

    def __init__(self, parent, controller):
        super().__init__(parent, bg="white")
        self.controller = controller
        self._images = []  # mantém referência das imagens (PhotoImage) vivas

        scroll = ScrollableFrame(self, bg="white")
        scroll.pack(fill="both", expand=True)
        self.body = scroll.inner

        self._build_content()

    # ------------------------------------------------------------------
    # Helpers de layout
    # ------------------------------------------------------------------
    def _wrapper(self):
        f = tk.Frame(self.body, bg="white")
        f.pack(fill="x", padx=40, pady=(0, 4))
        return f

    def _add_title(self, text):
        f = self._wrapper()
        tk.Label(f, text=text, font=("Segoe UI", 20, "bold"), bg="white", fg="#111111",
                 anchor="w", justify="left", wraplength=self.CONTENT_WIDTH).pack(anchor="w", pady=(28, 4))

    def _add_intro(self, text):
        f = self._wrapper()
        tk.Label(f, text=text, font=("Segoe UI", 10), bg="white", fg=self.MUTED_COLOR,
                 anchor="w", justify="left", wraplength=self.CONTENT_WIDTH).pack(anchor="w", pady=(0, 10))

    def _add_section(self, text):
        f = self._wrapper()
        border = tk.Frame(f, bg=self.ACCENT_COLOR, height=3)
        border.pack(fill="x", pady=(18, 8))
        tk.Label(f, text=text, font=("Segoe UI", 15, "bold"), bg="white", fg="#111111",
                 anchor="w", justify="left", wraplength=self.CONTENT_WIDTH).pack(anchor="w", pady=(0, 6))

    def _add_subheader(self, text):
        f = self._wrapper()
        tk.Label(f, text=text, font=("Segoe UI", 12, "bold"), bg="white", fg=self.ACCENT_COLOR,
                 anchor="w", justify="left", wraplength=self.CONTENT_WIDTH).pack(anchor="w", pady=(14, 4))

    def _add_paragraph(self, text):
        f = self._wrapper()
        tk.Label(f, text=text, font=("Segoe UI", 10), bg="white", fg=self.TEXT_COLOR,
                 anchor="w", justify="left", wraplength=self.CONTENT_WIDTH).pack(anchor="w", pady=(0, 8))

    def _add_bullets(self, items):
        f = self._wrapper()
        for item in items:
            row = tk.Frame(f, bg="white")
            row.pack(fill="x", anchor="w", pady=2)
            tk.Label(row, text="•", font=("Segoe UI", 10, "bold"), bg="white", fg=self.ACCENT_COLOR).pack(side="left", anchor="n", padx=(0, 8))
            tk.Label(row, text=item, font=("Segoe UI", 10), bg="white", fg=self.TEXT_COLOR,
                     anchor="w", justify="left", wraplength=self.CONTENT_WIDTH - 20).pack(side="left", anchor="w", fill="x", expand=True)

    def _add_warning(self, text):
        f = self._wrapper()
        box = tk.Frame(f, bg=self.WARNING_BG, highlightbackground=self.WARNING_BORDER, highlightthickness=1)
        box.pack(fill="x", pady=(4, 10))
        tk.Label(box, text=text, font=("Segoe UI", 10, "bold"), bg=self.WARNING_BG, fg="#713f12",
                 anchor="w", justify="left", wraplength=self.CONTENT_WIDTH - 24).pack(anchor="w", padx=12, pady=10)

    def _add_link(self, text, url):
        f = self._wrapper()
        lbl = tk.Label(f, text=text, font=("Segoe UI", 10, "underline"), bg="white", fg=self.ACCENT_COLOR,
                        anchor="w", cursor="hand2")
        lbl.pack(anchor="w", pady=(0, 8))
        lbl.bind("<Button-1>", lambda e: webbrowser.open(url))

    def _add_image(self, filename, caption=None):
        f = self._wrapper()
        img_box = tk.Frame(f, bg="#e2e8f0", highlightbackground="#cbd5e1", highlightthickness=1)
        img_box.pack(anchor="w", pady=(4, 4))

        if _PIL_OK:
            try:
                path = resource_path(os.path.join("assets", "tutorial", filename))
                pil_img = Image.open(path)
                w, h = pil_img.size
                if w > self.CONTENT_WIDTH:
                    ratio = self.CONTENT_WIDTH / w
                    new_size = (self.CONTENT_WIDTH, max(1, int(h * ratio)))
                    if _RESAMPLE is not None:
                        pil_img = pil_img.resize(new_size, _RESAMPLE)
                    else:
                        pil_img = pil_img.resize(new_size)
                photo = ImageTk.PhotoImage(pil_img)
                self._images.append(photo)
                tk.Label(img_box, image=photo, bg="#e2e8f0").pack()
            except Exception:
                tk.Label(img_box, text=f"[Imagem indisponível: {filename}]", font=("Segoe UI", 9),
                         bg="#e2e8f0", fg="#64748b", padx=20, pady=20).pack()
        else:
            tk.Label(img_box, text="[Instale a biblioteca Pillow para ver as imagens do tutorial]",
                     font=("Segoe UI", 9), bg="#e2e8f0", fg="#64748b", padx=20, pady=20).pack()

        if caption:
            tk.Label(f, text=caption, font=("Segoe UI", 9, "italic"), bg="white", fg="#64748b",
                     anchor="w", justify="left", wraplength=self.CONTENT_WIDTH).pack(anchor="w", pady=(2, 10))
        else:
            tk.Frame(f, bg="white", height=10).pack()

    def _add_spacer(self, height=16):
        tk.Frame(self.body, bg="white", height=height).pack()

    # ------------------------------------------------------------------
    # Conteúdo do tutorial
    # ------------------------------------------------------------------
    def _build_content(self):
        self._add_title("Como utilizar o SIFSTAT")
        self._add_intro(
            "Siga o passo a passo abaixo para exportar os mapas do ERP e do PGA-SIGSIF "
            "e depois compará-los automaticamente no SIFSTAT."
        )

        # ============================================================
        # 1º - Exportar os mapas
        # ============================================================
        self._add_section("1º Vamos exportar os Mapas que faremos a conferência")

        self._add_subheader("No ERP")
        self._add_bullets([
            "Na tela \"Mapa de Movimentação de Estoque\", na base filial, selecione o tipo de mapa que vai "
            "conferir: Produção, Recebimento ou Comercialização.",
            "Selecione o SIF da sua unidade.",
            "Filtre o mês referente (tem que ser o mês fechado).",
            "Em \"Exibe Produto\", selecione Não.",
            "Clique em Executar.",
        ])
        self._add_image("passo_01_erp_parametros.png")

        self._add_paragraph("Abrirá uma aba com o resultado.")
        self._add_warning(
            "MUITO IMPORTANTE! Exclua o filtro de \"Data\" que aparece logo acima dos dados gerados. "
            "Se não excluir, o programa não vai funcionar corretamente.\n\n"
            "Para excluir, segure em cima do filtro e arraste para cima - vai aparecer um \"X\", é só soltar. "
            "Veja em amarelo, na imagem abaixo, o filtro que precisa ser excluído:"
        )
        self._add_image("passo_02_erp_resultado_filtro_data.png")

        self._add_paragraph(
            "Agora, exporte em \"Excel\" - o sistema só funciona se o arquivo estiver nesse formato. "
            "Salve a cópia com um nome que ajude a identificar. Exemplo: MAPA PRODUÇÃO 08-2026."
        )
        self._add_image("passo_03_erp_exportar_excel.png")

        self._add_paragraph("Pronto, exportamos o mapa do ERP.")

        self._add_subheader("No PGA")
        self._add_paragraph("Acesse o site do PGA-SIGSIF:")
        self._add_link("https://sistemas.agricultura.gov.br/pga_sigsif/", "https://sistemas.agricultura.gov.br/pga_sigsif/")

        self._add_paragraph(
            "Selecione a aba \"Processo\" e escolha a opção \"Mapas Estatísticos\". Escolha o mesmo tipo de "
            "mapa que você exportou no ERP (Produção, Recebimento ou Comercialização)."
        )
        self._add_image("passo_04_pga_mapas_estatisticos.png")

        self._add_paragraph(
            "Preencha os dados da empresa e o mês referente. Clique em \"Consultar\" e depois em "
            "\"Imprimir Relatório\"."
        )
        self._add_image("passo_05_pga_consultar_imprimir.png")

        self._add_paragraph(
            "Vai abrir uma tela acima - filtre novamente o mês fechado, clique em \"Consultar\" e depois em "
            "\"Exportar XLS\". Salve novamente com um nome que ajude a identificar facilmente."
        )
        self._add_image("passo_06_pga_exportar_xls.png")

        # ============================================================
        # 2º - Comparar no SIFSTAT
        # ============================================================
        self._add_section("2º Vamos realizar a comparação dos Mapas no SIFSTAT")

        self._add_paragraph(
            "Abra o SIFSTAT, escolha a opção \"Mapas Estatísticos\" e o tipo de mapa referente."
        )
        self._add_image("passo_07_sifstat_menu.png")

        self._add_paragraph(
            "Selecione agora o mapa exportado do ERP e o do PGA e aperte em \"Executar\"."
        )
        self._add_image("passo_08_sifstat_selecionar_arquivos.png")

        self._add_paragraph(
            "O sistema irá realizar a comparação e mostrar caso haja algum produto com diferença no ERP/PGA."
        )
        self._add_image("passo_09_sifstat_resultado.png")

        self._add_paragraph(
            "Através da opção \"Mostrar apenas divergências\", o sistema separa apenas os produtos que "
            "estiverem com diferença entre ERP/PGA.\n\n"
            "Na opção \"Exportar Excel\", o sistema exporta uma planilha com as conferências realizadas."
        )

        self._add_paragraph("E assim funciona o SIFSTAT!\nPara dúvidas, utilize a aba \"Contato\".")

        self._add_spacer(30)


class ContatoPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="white")
        self.controller = controller

        wrapper = tk.Frame(self, bg="white")
        wrapper.pack(anchor="nw", padx=50, pady=50)

        tk.Label(wrapper, text="Contato", font=("Segoe UI", 20, "bold"), bg="white", fg="#111111").pack(anchor="w", pady=(0, 20))

        tk.Label(
            wrapper,
            text="Caso haja problemas ou dificuldades quando for utilizar.",
            font=("Segoe UI", 11), bg="white", fg="#1e293b", justify="left"
        ).pack(anchor="w", pady=(0, 4))

        tk.Label(
            wrapper,
            text="Envie e-mail para:",
            font=("Segoe UI", 11), bg="white", fg="#1e293b", justify="left"
        ).pack(anchor="w", pady=(0, 10))

        tk.Label(
            wrapper,
            text="victor.vieira@friboi.com.br",
            font=("Segoe UI", 12, "bold"), bg="white", fg="#2563eb", justify="left"
        ).pack(anchor="w")

        tk.Label(
            wrapper,
            text="heitor.lopes@friboi.com.br",
            font=("Segoe UI", 12, "bold"), bg="white", fg="#2563eb", justify="left"
        ).pack(anchor="w", pady=(2, 0))


# ==========================================================================
# MENU LATERAL (SIDEBAR)
# ==========================================================================

class Sidebar(tk.Frame):
    SIDEBAR_BG = "#fafafa"
    HOVER_BG = "#eeeeee"
    ACTIVE_BG = "#e5e5e5"

    def __init__(self, parent, controller):
        super().__init__(parent, bg=self.SIDEBAR_BG, width=210)
        self.controller = controller
        self.pack_propagate(False)

        self.nav_buttons = {}       # nome_pagina -> widget do botão
        self.submenu_visible = False

        self._build()

    def _build(self):
        # Título
        tk.Label(
            self, text="SIFSTAT", font=("Segoe UI", 15, "bold"),
            bg=self.SIDEBAR_BG, fg="#111111", anchor="w"
        ).pack(fill="x", padx=18, pady=(20, 22))

        # --- Mapas Estatísticos (expande submenu) ---
        self.btn_mapas = tk.Button(
            self, text="Mapas Estatísticos", font=("Segoe UI", 10, "bold"),
            bg=self.SIDEBAR_BG, fg="#111111", bd=0, anchor="w",
            padx=18, pady=9, activebackground=self.HOVER_BG, activeforeground="#111111",
            cursor="hand2", command=self.toggle_submenu
        )
        self.btn_mapas.pack(fill="x")

        # Submenu (Produção / Recebimento / Comercialização)
        self.submenu_frame = tk.Frame(self, bg=self.SIDEBAR_BG)
        self._add_nav_button(self.submenu_frame, "Produção", "producao", indent=True)
        self._add_nav_button(self.submenu_frame, "Recebimento", "recebimento", indent=True)
        self._add_nav_button(self.submenu_frame, "Comercialização", "comercializacao", indent=True)

        # --- Como utilizar ? ---
        self._add_nav_button(self, "Como utilizar ?", "como_utilizar")

        # --- Contato ---
        self._add_nav_button(self, "Contato", "contato")

        # --- Rodapé: assinatura + logo JBS ---
        bottom = tk.Frame(self, bg=self.SIDEBAR_BG)
        bottom.pack(side="bottom", fill="x", pady=(0, 18))

        tk.Label(
            bottom,
            text="Desenvolvido por:\nVíctor Sanches - FATURAMENTO JBS-GYN\nHeitor Lopes - CSC FISCAL JBS-GYN",
            font=("Segoe UI", 7), bg=self.SIDEBAR_BG, fg="#555555", justify="center"
        ).pack(pady=(0, 10))

        try:
            self._logo_img = tk.PhotoImage(file=resource_path(os.path.join("assets", "jbs_logo.png")))
            tk.Label(bottom, image=self._logo_img, bg=self.SIDEBAR_BG).pack()
        except Exception:
            pass

    def _add_nav_button(self, parent, text, page_name, indent=False):
        btn = tk.Button(
            parent, text=text, font=("Segoe UI", 10),
            bg=self.SIDEBAR_BG, fg="#111111", bd=0, anchor="w",
            padx=(34 if indent else 18), pady=8,
            activebackground=self.HOVER_BG, activeforeground="#111111",
            cursor="hand2", command=lambda: self.controller.show_page(page_name)
        )
        btn.pack(fill="x")
        self.nav_buttons[page_name] = btn
        return btn

    def toggle_submenu(self):
        if self.submenu_visible:
            self.submenu_frame.pack_forget()
        else:
            self.submenu_frame.pack(fill="x", after=self.btn_mapas)
        self.submenu_visible = not self.submenu_visible

    def set_active(self, page_name):
        """Realça no menu a página atualmente aberta."""
        for name, btn in self.nav_buttons.items():
            btn.config(bg=self.ACTIVE_BG if name == page_name else self.SIDEBAR_BG)

        # Se uma das páginas do submenu está ativa, garante que o submenu esteja visível
        if page_name in ("producao", "recebimento", "comercializacao") and not self.submenu_visible:
            self.submenu_frame.pack(fill="x", after=self.btn_mapas)
            self.submenu_visible = True


# ==========================================================================
# APLICATIVO PRINCIPAL
# ==========================================================================

class SifstatApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SIFSTAT")
        self.geometry("1300x820")
        self.minsize(1000, 650)
        self.configure(bg="white")

        try:
            self.iconbitmap(resource_path(os.path.join("assets", "icone.ico")))
        except Exception:
            pass

        container = tk.Frame(self, bg="white")
        container.pack(fill="both", expand=True)

        self.sidebar = Sidebar(container, self)
        self.sidebar.pack(side="left", fill="y")

        separator = tk.Frame(container, bg="#e2e2e2", width=1)
        separator.pack(side="left", fill="y")

        self.content = tk.Frame(container, bg="white")
        self.content.pack(side="left", fill="both", expand=True)
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        self.pages = {}
        for name, PageClass in (
            ("home", HomePage),
            ("producao", ProducaoFrame),
            ("recebimento", RecebimentoFrame),
            ("comercializacao", ComercializacaoFrame),
            ("como_utilizar", ComoUtilizarPage),
            ("contato", ContatoPage),
        ):
            page = PageClass(self.content, self)
            page.grid(row=0, column=0, sticky="nsew")
            self.pages[name] = page

        self.show_page("home")

    def show_page(self, name):
        page = self.pages.get(name)
        if page is None:
            return
        page.tkraise()
        self.sidebar.set_active(name)


if __name__ == "__main__":
    try:
        import ctypes
        myappid = "jbs.sifstat.unificado.v1"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    app = SifstatApp()
    app.mainloop()
