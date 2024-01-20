import csv
import os
from tkinter import *
from tkinter import ttk, simpledialog, messagebox
from ttkthemes import ThemedTk

class CadastroAlunosCursos:
    def __init__(self):
        self.root = ThemedTk(theme="arc")  # Escolha um tema, por exemplo, "arc"
        self.root.title("Cadastro de Alunos e Cursos - IJCPM")

        self.tree = ttk.Treeview(self.root, columns=(1, 2, 3, 4, 5, 6, 7, 8, 9, 10), show="headings", height="5")
        self.tree.heading(1, text="Nome")
        self.tree.heading(2, text="Idade")
        self.tree.heading(3, text="Sexo")
        self.tree.heading(4, text="Endereço")
        self.tree.heading(5, text="E-mail")
        self.tree.heading(6, text="Telefone")
        self.tree.heading(7, text="CPF")
        self.tree.heading(8, text="Nome do Curso")
        self.tree.heading(9, text="Prazo")
        self.tree.heading(10, text="Detalhes do Curso")
        self.tree.pack(padx=10, pady=10)

        btn_cadastrar_aluno_curso = Button(self.root, text="Cadastrar Aluno e Curso", command=self.cadastrar_aluno_curso)
        btn_cadastrar_aluno_curso.pack(pady=5)

        btn_listar_alunos_cursos = Button(self.root, text="Listar Alunos e Cursos Cadastrados", command=self.listar_alunos_cursos)
        btn_listar_alunos_cursos.pack(pady=5)

        btn_excluir_aluno_curso = Button(self.root, text="Excluir Aluno e Curso", command=self.excluir_aluno_curso)
        btn_excluir_aluno_curso.pack(pady=5)

        btn_detalhes_aluno_curso = Button(self.root, text="Detalhes do Aluno e Curso", command=self.detalhes_aluno_curso)
        btn_detalhes_aluno_curso.pack(pady=5)

        self.label_status = Label(self.root, text="", fg="green")
        self.label_status.pack(pady=10)

        # Adiciona um evento de clique duplo na Treeview
        self.tree.bind("<Double-1>", self.visualizar_detalhes)

        self.root.mainloop()

    def cadastrar_aluno_curso(self):
        nome = simpledialog.askstring("Cadastro", "Digite o nome completo do aluno:")
        if nome is None:
            return  # O usuário clicou em Cancelar

        idade = simpledialog.askinteger("Cadastro", "Digite a idade do aluno:")
        if idade is None:
            return

        sexo = simpledialog.askstring("Cadastro", "Digite o sexo do aluno:")
        if sexo is None:
            return

        endereco = simpledialog.askstring("Cadastro", "Digite o endereço do aluno:")
        if endereco is None:
            return

        email = simpledialog.askstring("Cadastro", "Digite o e-mail do aluno:")
        if email is None:
            return

        telefone = simpledialog.askstring("Cadastro", "Digite o telefone do aluno:")
        if telefone is None:
            return

        cpf = simpledialog.askstring("Cadastro", "Digite o CPF do aluno:")
        if cpf is None:
            return

        nome_curso = simpledialog.askstring("Cadastro", "Digite o nome do curso:")
        if nome_curso is None:
            return

        prazo = simpledialog.askinteger("Cadastro", "Digite o prazo do curso (em meses):")
        if prazo is None:
            return

        inicio_curso = simpledialog.askstring("Cadastro", "Digite a data de início do curso (DD/MM/AAAA):")
        if inicio_curso is None:
            return

        fim_curso = simpledialog.askstring("Cadastro", "Digite a data de término do curso (DD/MM/AAAA):")
        if fim_curso is None:
            return

        detalhes_curso = simpledialog.askstring("Cadastro", "Digite detalhes adicionais do curso:")
        if detalhes_curso is None:
            return

        dados_aluno_curso = [nome, idade, sexo, endereco, email, telefone, cpf, nome_curso, prazo, detalhes_curso]

        with open('alunos_cursos.csv', mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(dados_aluno_curso)

        self.label_status.config(text="Cadastro do aluno e curso realizado com sucesso!", fg="green")
        self.listar_alunos_cursos()

    def listar_alunos_cursos(self):
        self.tree.delete(*self.tree.get_children())  # Limpar a Treeview antes de exibir novamente
        if os.path.exists('alunos_cursos.csv'):
            with open('alunos_cursos.csv', mode='r') as file:
                reader = csv.reader(file)
                for row in reader:
                    self.tree.insert("", "end", values=row)

            self.label_status.config(text="")
        else:
            self.label_status.config(text="Nenhum aluno e curso cadastrado.", fg="red")

    def excluir_aluno_curso(self):
        selected_item = self.tree.selection()
        if not selected_item:
            self.label_status.config(text="Selecione um aluno e curso para excluir.", fg="red")
            return

        confirmed = messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir este aluno e curso?")
        if confirmed:
            selected_id = self.tree.item(selected_item, 'values')[0]

            with open('alunos_cursos.csv', 'r') as file:
                data = list(csv.reader(file))

            # Encontrar o índice do aluno e curso com base no Nome do Aluno
            index_to_remove = None
            for i, row in enumerate(data):
                if row and row[0] == selected_id:
                    index_to_remove = i
                    break

            if index_to_remove is not None:
                del data[index_to_remove]

                with open('alunos_cursos.csv', 'w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerows(data)

                self.label_status.config(text="Aluno e curso excluídos com sucesso!", fg="green")
                self.listar_alunos_cursos()
            else:
                self.label_status.config(text="Nome do aluno não encontrado.", fg="red")

    def visualizar_detalhes(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            selected_id = self.tree.item(selected_item, 'values')[0]
            messagebox.showinfo("Detalhes do Aluno e Curso", f"Nome do Aluno: {selected_id}\nMais detalhes aqui...")

    def detalhes_aluno_curso(self):
        selected_item = self.tree.selection()
        if selected_item:
            selected_id = self.tree.item(selected_item, 'values')[0]
            messagebox.showinfo("Detalhes do Aluno e Curso", f"Nome do Aluno: {selected_id}\nMais detalhes aqui...")

if __name__ == "__main__":
    cadastro_alunos_cursos = CadastroAlunosCursos()
///////VERSÃO ATUALIZADA///////

import csv
from tkinter import *
from tkinter import ttk, simpledialog, messagebox
from ttkthemes import ThemedTk
import os
class CadastroAlunosCursos:
    def __init__(self):
        self.root = ThemedTk(theme="arc")  # Escolha um tema, por exemplo, "arc"
        self.root.title("Cadastro de Alunos e Cursos - IJCPM")

        self.inicializar_interface()

    def inicializar_interface(self):
        self.tree = ttk.Treeview(self.root, columns=(1, 2, 3, 4, 5, 6, 7, 8, 9, 10), show="headings", height="5")
        self.tree.heading(1, text="Nome")
        self.tree.heading(2, text="Idade")
        self.tree.heading(3, text="Sexo")
        self.tree.heading(4, text="Endereço")
        self.tree.heading(5, text="E-mail")
        self.tree.heading(6, text="Telefone")
        self.tree.heading(7, text="CPF")
        self.tree.heading(8, text="Nome do Curso")
        self.tree.heading(9, text="Prazo")
        self.tree.heading(10, text="Detalhes do Curso")
        self.tree.pack(padx=10, pady=10)

        btn_cadastrar_aluno_curso = Button(self.root, text="Cadastrar Aluno e Curso", command=self.cadastrar_aluno_curso)
        btn_cadastrar_aluno_curso.pack(pady=5)

        btn_listar_alunos_cursos = Button(self.root, text="Listar Alunos e Cursos Cadastrados", command=self.listar_alunos_cursos)
        btn_listar_alunos_cursos.pack(pady=5)

        btn_editar_aluno_curso = Button(self.root, text="Editar Aluno e Curso", command=self.editar_aluno_curso)
        btn_editar_aluno_curso.pack(pady=5)

        btn_excluir_aluno_curso = Button(self.root, text="Excluir Aluno e Curso", command=self.excluir_aluno_curso)
        btn_excluir_aluno_curso.pack(pady=5)

        btn_detalhes_aluno_curso = Button(self.root, text="Detalhes do Aluno e Curso", command=self.detalhes_aluno_curso)
        btn_detalhes_aluno_curso.pack(pady=5)

        self.label_status = Label(self.root, text="", fg="green")
        self.label_status.pack(pady=10)

        # Adiciona um evento de clique duplo na Treeview
        self.tree.bind("<Double-1>", self.visualizar_detalhes)

        self.root.mainloop()

    def cadastrar_aluno_curso(self):
        campos = [
            ("Nome do Aluno", "Digite o nome completo do aluno:"),
            ("Idade do Aluno", "Digite a idade do aluno:"),
            ("Sexo do Aluno", "Digite o sexo do aluno:"),
            ("Endereço do Aluno", "Digite o endereço do aluno:"),
            ("E-mail do Aluno", "Digite o e-mail do aluno:"),
            ("Telefone do Aluno", "Digite o telefone do aluno:"),
            ("CPF do Aluno", "Digite o CPF do aluno:"),
            ("Nome do Curso", "Digite o nome do curso:"),
            ("Prazo do Curso (em meses)", "Digite o prazo do curso:"),
            ("Detalhes do Curso", "Digite detalhes adicionais do curso:")
        ]

        dados_aluno_curso = []

        for label, prompt in campos:
            resposta = simpledialog.askstring("Cadastro", prompt)
            if resposta is None:
                return
            dados_aluno_curso.append(resposta)

        with open('alunos_cursos.csv', mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(dados_aluno_curso)

        self.label_status.config(text="Cadastro do aluno e curso realizado com sucesso!", fg="green")
        self.listar_alunos_cursos()

    def listar_alunos_cursos(self):
        self.tree.delete(*self.tree.get_children())  # Limpar a Treeview antes de exibir novamente
        if os.path.exists('alunos_cursos.csv'):
            with open('alunos_cursos.csv', mode='r') as file:
                reader = csv.reader(file)
                for row in reader:
                    self.tree.insert("", "end", values=row)

            self.label_status.config(text="")
        else:
            self.label_status.config(text="Nenhum aluno e curso cadastrado.", fg="red")

    def excluir_aluno_curso(self):
        selected_item = self.tree.selection()
        if not selected_item:
            self.label_status.config(text="Selecione um aluno e curso para excluir.", fg="red")
            return

        confirmed = messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir este aluno e curso?")
        if confirmed:
            selected_id = self.tree.item(selected_item, 'values')[0]

            with open('alunos_cursos.csv', 'r') as file:
                data = list(csv.reader(file))

            # Encontrar o índice do aluno e curso com base no Nome do Aluno
            index_to_remove = None
            for i, row in enumerate(data):
                if row and row[0] == selected_id:
                    index_to_remove = i
                    break

            if index_to_remove is not None:
                del data[index_to_remove]

                with open('alunos_cursos.csv', 'w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerows(data)

                self.label_status.config(text="Aluno e curso excluídos com sucesso!", fg="green")
                self.listar_alunos_cursos()
            else:
                self.label_status.config(text="Nome do aluno não encontrado.", fg="red")

    def visualizar_detalhes(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            selected_id = self.tree.item(selected_item, 'values')[0]
            messagebox.showinfo("Detalhes do Aluno e Curso", f"Nome do Aluno: {selected_id}\nMais detalhes aqui...")

    def detalhes_aluno_curso(self):
        selected_item = self.tree.selection()
        if selected_item:
            selected_id = self.tree.item(selected_item, 'values')[0]
            messagebox.showinfo("Detalhes do Aluno e Curso", f"Nome do Aluno: {selected_id}\nMais detalhes aqui...")

    def editar_aluno_curso(self):
        selected_item = self.tree.selection()
        if selected_item:
            # Recupera os dados atuais do aluno e curso
            dados_atuais = list(self.tree.item(selected_item, 'values'))

            campos = [
                ("Nome do Aluno", dados_atuais[0]),
                ("Idade do Aluno", dados_atuais[1]),
                ("Sexo do Aluno", dados_atuais[2]),
                ("Endereço do Aluno", dados_atuais[3]),
                ("E-mail do Aluno", dados_atuais[4]),
                ("Telefone do Aluno", dados_atuais[5]),
                ("CPF do Aluno", dados_atuais[6]),
                ("Nome do Curso", dados_atuais[7]),
                ("Prazo do Curso", dados_atuais[8]),
                ("Detalhes do Curso", dados_atuais[9])
            ]

            novos_dados = []

            for label, valor_atual in campos:
                resposta = simpledialog.askstring("Edição", f"{label} (atual: {valor_atual}):")
                if resposta is None:
                    return
                novos_dados.append(resposta)

            # Remove o aluno e curso atual
            self.excluir_aluno_curso()

            # Adiciona os novos dados
            with open('alunos_cursos.csv', mode='a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(novos_dados)

            self.label_status.config(text="Edição do aluno e curso realizada com sucesso!", fg="green")
            self.listar_alunos_cursos()

if __name__ == "__main__":
    cadastro_alunos_cursos = CadastroAlunosCursos()
