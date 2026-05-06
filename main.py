import time
import threading
import sys
import keyboard
import pandas as pd
import pyautogui as py
from datetime import datetime
from tkcalendar import DateEntry
from tkinter import *
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, msg):
        self.widget.after(0, self._append, msg)

    def _append(self, msg):
        self.widget.insert("end", msg)
        self.widget.see("end")

        self.widget.update_idletasks()

    def flush(self):
        pass

class Application:

    def centralizar_janela(self, janela):
        janela.update_idletasks()

        largura = janela.winfo_width()
        altura = janela.winfo_height()

        largura_tela = janela.winfo_screenwidth()
        altura_tela = janela.winfo_screenheight()

        x = (largura_tela // 2) - (largura // 2)
        y = (altura_tela // 2) - (altura // 2)

        janela.geometry(f"+{x}+{y}")

    def selecionar_arquivo(self, entry):

        caminho = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=[
                ("Arquivos Excel", "*.xls *.xlsx"),
                ("Todos os arquivos", "*.*")
            ]
        )

        if caminho:
            entry.delete(0, END)
            entry.insert(0, caminho)

    def executar_thread(self, comando):
        threading.Thread(target=comando, daemon=True).start()

    def __init__(self, master=None):

        self.parar = False
        self.executando = False
        keyboard.add_hotkey('esc', self.forcar_parada)

        self.master = master

        self.fontePadrao = ("Arial", "10")
        self.fontePadraoBold = ("Arial", "10", "bold")

        main = Frame(master)
        main.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

        container = Frame(main, width=400, height=500)
        container.grid(row=0, column=0)
        container.grid_propagate(False)

        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        master.grid_rowconfigure(0, weight=1)
        master.grid_columnconfigure(0, weight=1)

        header = Frame(container)
        header.grid(row=0, column=0, pady=(0, 15))

        header.grid_columnconfigure(0, weight=1)

        titulo = Label(
            header, text="Lançar Apontamentos", font=self.fontePadraoBold
        )
        titulo.grid(row=0, column=1, sticky="e")

        configBtn = Button(
            header,
            text="⚙"
            )
        configBtn.grid(row=0, column=2)

        forms = Frame(container)
        forms.grid(row=1, column=0, pady=10)

        forms.grid_columnconfigure(1, weight=1)

        planilhaLabel = Label(
            forms,
            text="Planilha:",
            font=self.fontePadrao
            )
        
        planilhaLabel.grid(row=0, column=0, sticky="e", pady=5, padx=5)

        self.planilhaEntry = Entry(
            forms,
            font=self.fontePadrao,
            width=30
            )

        self.planilhaEntry.grid(row=0, column=1, sticky="w", pady=5)

        planilhaBtn = Button(
            forms,
            text="📁",
            command=lambda: self.selecionar_arquivo(self.planilhaEntry)
            )

        planilhaBtn.grid(row=0, column=2, sticky="w", padx=5)

        dataInicialLabel = Label(
            forms,
            text="Data Inicial:",
            font=self.fontePadrao
            )
        dataInicialLabel.grid(row=1, column=0, sticky="e", pady=5, padx=5)

        self.dataInicialEntry = DateEntry(
            forms,
            width=12,
            font=self.fontePadrao,
            locale="pt_br",
            date_pattern="dd/mm/yyyy"
            )
        self.dataInicialEntry.grid(row=1, column=1, sticky="w", pady=5)

        dataFinalLabel = Label(
            forms,
            text="Data Final:",
            font=self.fontePadrao
            )
        dataFinalLabel.grid(row=2, column=0, sticky="e", pady=5, padx=5)

        self.dataFinalEntry = DateEntry(
            forms,
            width=12,
            font=self.fontePadrao,
            locale="pt_br",
            date_pattern="dd/mm/yyyy"
            )
        self.dataFinalEntry.grid(row=2, column=1, sticky="w", pady=5)

        self.executarBtn = Button(
            container,
            text="Executar",
            font=self.fontePadraoBold,
            command= lambda: self.executar_thread(self.executar),
            width = 20
        )
        self.executarBtn.grid(row=2, column=0, pady=15)

        progress_frame = Frame(container)
        progress_frame.grid(row=3, column=0, sticky="ew", pady=(5,10))

        self.progressLabel = Label(
            progress_frame,
            text="",
            font=self.fontePadrao
        )
        self.progressLabel.pack(anchor="w")

        self.progressBar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
        )

        self.progressBar.pack(fill="x", expand=True)

        console_frame = Frame(container)
        console_frame.grid(row=4, column=0, sticky="nsew")

        main.grid_rowconfigure(4, weight=1)

        self.console = scrolledtext.ScrolledText(
            console_frame,
            state="normal",
            height=15,
            bg="black",
            fg="white"
        )

        self.console.pack(fill="both", expand=True)
        sys.stdout = TextRedirector(self.console)

        self.centralizar_janela(master)

    def forcar_parada(self):
        if self.executando:
            self.parar = True
            self.executando = False
            py.FAILSAFE = True
            py.moveTo(0, 0)
            messagebox.showerror(
                "Execução Interrompida", "O lançamento foi interrompido!\nO usuário forçou a parada."
            )
            print("Parado pelo usuário <<ESC>>")

    def validar_entrada(self):

        arquivo = self.planilhaEntry.get()
        dataInicial = self.dataInicialEntry.get()
        dataFinal = self.dataFinalEntry.get()

        if not arquivo:
            messagebox.showerror(
                "Erro", "Por favor, selecione uma planilha!"
            )
            return False
        
        if not dataInicial:
            messagebox.showerror(
                "Erro", "Por favor, insira a data de início!"
            )
            return False
        
        if not dataFinal:
            messagebox.showerror(
                "Erro", "Por favor, insira a data de fim!"
            )
            return False
        return True

    def executar(self):

        print("Executando...")
        self.executando = True

        if not self.validar_entrada():
            print("Validação falhou. Corrija os erros e tente novamente.")
            return

        py.FAILSAFE = True
        py.PAUSE = 0.1
        self.parar = False
        TEMPO = 1.2
        CAMINHO = self.planilhaEntry.get()
        DATA_INICIO = self.dataInicialEntry.get_date()
        DATA_FIM = self.dataFinalEntry.get_date()

        mes_atual = datetime.now().month
        mes_inicio = DATA_INICIO.month
        mes_fim = DATA_FIM.month

        print(f"Executando com: {CAMINHO}, {DATA_INICIO}, {DATA_FIM}\n")

        def check_stop():
            if self.parar:
                print("Ação interrompida pelo usuário.")
                self.executando = False
                return True
            return False

        def esperar():
            for _ in range(int(TEMPO * 10)):
                if self.parar:
                    self.executando = False
                    return
                time.sleep(0.1)

        def abrir_menu():

            if check_stop(): return
            py.click(110, 37)
            if check_stop(): return
            py.click(110, 63)
            if check_stop(): return
            py.click(343, 129)
            if check_stop(): return
            py.click(515, 129)
            esperar()

        def selecionar_funcionario(codigo):
            
            if check_stop(): return
            py.click(428, 211) 
            esperar()

            if check_stop(): return
            py.click(524, 217)
            if check_stop(): return
            py.write(str(int(codigo)))
            esperar()

            if check_stop(): return
            py.press('enter')
            esperar()

        def novo():
            if check_stop(): return
            py.click(410, 238)
            esperar()

        def evento(cod):
            if check_stop(): return
            py.click(400, 290)
            if check_stop(): return
            py.write(str(int(cod)))
            if check_stop(): return
            py.press('enter')
            esperar()

        def parametro(valor):
            if check_stop(): return
            py.click(400, 330)
            if check_stop(): return
            py.write(str(valor).replace(",", "."))
            esperar()

        def datas():
            if check_stop(): return
            py.click(615, 330)

            if check_stop(): return
            if mes_atual == mes_inicio:
                py.press('1')
            elif mes_atual > mes_inicio:
                for _ in range(mes_atual - mes_inicio):
                    if check_stop(): return
                    py.press('down')
                    esperar()
            elif mes_atual < mes_inicio:
                for _ in range(mes_inicio - mes_atual):
                    if check_stop(): return
                    py.press('up')
                    esperar()

            if check_stop(): return
            py.press('tab')
            
            if check_stop(): return
            if mes_atual == mes_fim:
                py.press('1')
            elif mes_atual > mes_fim:
                for _ in range(mes_atual - mes_fim):
                    if check_stop(): return
                    py.press('down')
                    esperar()
            elif mes_atual < mes_fim:
                for _ in range(mes_fim - mes_atual):
                    if check_stop(): return
                    py.press('up')
                    esperar()

            esperar()

        def gravar():
            if check_stop(): return
            py.click(840, 330)
            esperar()

        df = pd.read_excel(CAMINHO, header=3)

        df = df[df['Nome'].notna()]
        df = df[df['Ref./Valor'].notna()]

        df = df.sort_values(by=['Código'])

        print("Abra o sistema... (6s)")
        time.sleep(6)

        abrir_menu()

        cod_atual = None
        nome_atual = None

        for i, row in df.iterrows():

            if self.parar:
                messagebox.showerror(
                    "Execução Interrompida", "O lançamento foi interrompido!"
                )
                print("Execução interrompida.")
                self.executando = False
                return

            try:
                codigo = row['Código']
                nome = row['Nome']
                evento_cod = row['Evento']
                valor = row['Ref./Valor']

                print(f"Funcionário {codigo} - Evento {evento_cod}")

                # só troca funcionário quando muda
                if codigo != cod_atual:

                    if cod_atual is not None:

                        resposta = messagebox.askyesno("Confirmação", f"Deseja continuar?\nOs lançamentos do funcionário ({cod_atual}) {nome_atual} foram finalizados.\nAproveite para corrigir os lançamentos se necessário.")

                        if not resposta:
                            messagebox.showerror(
                                "Execução Interrompida", "O lançamento foi interrompido!\nO usuário optou por não continuar os lançamentos."
                            )
                            print("Execução interrompida.")
                            self.executando = False
                            self.parar = True
                            py.moveTo(0, 0)
                            return
                        else:
                            print("Continuando com os próximos lançamentos...")
                            esperar()

                    selecionar_funcionario(codigo)
                    cod_atual = codigo
                    nome_atual = nome

                novo()
                evento(evento_cod)
                parametro(valor)
                datas()
                gravar()

            except py.FailSafeException:
                messagebox.showerror(
                    "Execução Interrompida", "O lançamento foi interrompido!"
                )
                print("Interrompido pelo FAILSAFE (mouse no canto)")
                self.executando = False
                return

            except Exception as e:
                print(f"Erro linha {i}: {e}")
                continue

        print("FINALIZADO")
        self.executando = False
        messagebox.showinfo("Concluído", "Lançamentos finalizados com sucesso!")

root = Tk()
root.title("Lançamento de Apontamentos")
Application(root)
root.mainloop()
