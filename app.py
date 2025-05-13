import csv
import os
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook

CSV_FILE = "alunos_cursos.csv"
BACKUP_FILE = "backup_alunos_cursos.csv"

class SistemaCursos:
    def __init__(self, root):
        self.root = root
        self.root.title("Cadastro de Alunos em Cursos")
        self.root.geometry("1100x600")
        self.root.resizable(False, False)

        self.campos = [
            "Nome Completo", "Idade", "Email", "Telefone", "Data de Nascimento",
            "Sexo", "CPF", "Curso Desejado", "Turno", "Data de Início"
        ]
        self.entradas = {}

        self.criar_csv_se_nao_existir()

        # Formulário
        frame_form = tk.LabelFrame(self.root, text="Cadastrar Novo Aluno")
        frame_form.pack(pady=10)

        for i, campo in enumerate(self.campos):
            tk.Label(frame_form, text=campo).grid(row=i, column=0, sticky="e", padx=5, pady=2)
            entrada = tk.Entry(frame_form, width=40)
            entrada.grid(row=i, column=1, padx=5, pady=2)
            self.entradas[campo] = entrada

        botoes_frame = tk.Frame(frame_form)
        botoes_frame.grid(row=len(self.campos), column=1, pady=10)

        tk.Button(botoes_frame, text="Cadastrar", width=15, command=self.cadastrar).grid(row=0, column=0, padx=5)
        tk.Button(botoes_frame, text="Limpar", width=15, command=self.limpar_formulario).grid(row=0, column=1, padx=5)

        # Área de busca
        frame_busca = tk.LabelFrame(self.root, text="Buscar Alunos")
        frame_busca.pack(pady=5)

        tk.Label(frame_busca, text="Buscar por Nome ou Curso:").pack(side="left", padx=5)
        self.entrada_busca = tk.Entry(frame_busca, width=40)
        self.entrada_busca.pack(side="left", padx=5)
        tk.Button(frame_busca, text="Buscar", command=self.buscar_alunos).pack(side="left", padx=5)
        tk.Button(frame_busca, text="Mostrar Todos", command=self.carregar_alunos).pack(side="left", padx=5)

        # Tabela
        self.tree = ttk.Treeview(self.root, columns=self.campos, show="headings", height=15)
        for campo in self.campos:
            self.tree.heading(campo, text=campo)
            self.tree.column(campo, width=100)
        self.tree.pack(pady=10, padx=10, fill="x")

        # Exportar
        frame_export = tk.Frame(self.root)
        frame_export.pack()
        tk.Button(frame_export, text="Exportar para Excel", command=self.exportar_excel).pack(pady=10)

        self.carregar_alunos()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def criar_csv_se_nao_existir(self):
        if not os.path.exists(CSV_FILE):
            with open(CSV_FILE, mode='w', newline='', encoding='latin1') as file:
                writer = csv.writer(file)
                writer.writerow(self.campos)

    def cpf_ja_cadastrado(self, cpf):
        if not os.path.exists(CSV_FILE):
            return False
        with open(CSV_FILE, mode='r', newline='', encoding='latin1') as file:
            reader = csv.reader(file)
            next(reader, None)
            for row in reader:
                if len(row) > 6 and row[6] == cpf:
                    return True
        return False

    def cadastrar(self):
        dados = []
        for campo in self.campos:
            valor = self.entradas[campo].get().strip()
            if not valor:
                messagebox.showwarning("Campo obrigatório", f"Preencha o campo '{campo}'.")
                return
            dados.append(valor)

        if self.cpf_ja_cadastrado(dados[6]):
            messagebox.showerror("Erro", "Este CPF já está cadastrado.")
            return

        try:
            with open(CSV_FILE, mode='a', newline='', encoding='latin1') as file:
                writer = csv.writer(file)
                writer.writerow(dados)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar dados: {e}")
            return

        self.limpar_formulario()
        self.carregar_alunos()
        messagebox.showinfo("Sucesso", "Aluno cadastrado com sucesso!")

    def limpar_formulario(self):
        for entrada in self.entradas.values():
            entrada.delete(0, tk.END)

    def carregar_alunos(self):
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(CSV_FILE):
            return
        with open(CSV_FILE, mode='r', newline='', encoding='latin1') as file:
            reader = csv.reader(file)
            next(reader, None)
            for row in reader:
                if row:
                    self.tree.insert("", "end", values=row)

    def buscar_alunos(self):
        termo = self.entrada_busca.get().lower()
        self.tree.delete(*self.tree.get_children())
        try:
            with open(CSV_FILE, mode='r', newline='', encoding='latin1') as file:
                reader = csv.reader(file)
                next(reader, None)
                for row in reader:
                    if termo in row[0].lower() or termo in row[7].lower():
                        self.tree.insert("", "end", values=row)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar dados: {e}")

    def exportar_excel(self):
        if not os.path.exists(CSV_FILE):
            messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
            return

        wb = Workbook()
        ws = wb.active
        ws.append(self.campos)

        with open(CSV_FILE, mode='r', newline='', encoding='latin1') as file:
            reader = csv.reader(file)
            next(reader, None)
            for row in reader:
                ws.append(row)

        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            wb.save(filepath)
            messagebox.showinfo("Exportação", "Dados exportados com sucesso!")

    def on_close(self):
        self.salvar_backup()
        self.root.destroy()

    def salvar_backup(self):
        if os.path.exists(CSV_FILE):
            shutil.copy(CSV_FILE, BACKUP_FILE)

# Executar programa
if __name__ == "__main__":
    root = tk.Tk()
    app = SistemaCursos(root)
    root.mainloop()
