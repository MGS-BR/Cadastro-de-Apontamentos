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

    def flush(self):
        return


class Application:

    def __init__(self, master=None):

        self.master = master

        self.executando = False

        self.stop_event = threading.Event()

        self.master.bind("<Escape>", lambda e: self.forcar_parada())

        try:

            keyboard.add_hotkey(
                "ctrl+alt+p",
                self.forcar_parada,
                suppress=False,
                trigger_on_release=True,
            )

            keyboard.add_hotkey(
                "esc", self.forcar_parada, suppress=False, trigger_on_release=True
            )

        except Exception:

            print("Hotkeys globais indisponíveis.")

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

        titulo = Label(header, text="Lançar Apontamentos", font=self.fontePadraoBold)
        titulo.grid(row=0, column=1, sticky="e")

        configBtn = Button(header, text="⚙", command=self.config)
        configBtn.grid(row=0, column=2)

        forms = Frame(container)
        forms.grid(row=1, column=0, pady=10)

        forms.grid_columnconfigure(1, weight=1)

        planilhaLabel = Label(forms, text="Planilha:", font=self.fontePadrao)

        planilhaLabel.grid(row=0, column=0, sticky="e", pady=5, padx=5)

        self.planilhaEntry = Entry(forms, font=self.fontePadrao, width=30)

        self.planilhaEntry.grid(row=0, column=1, sticky="w", pady=5)

        planilhaBtn = Button(
            forms,
            text="📁",
            command=lambda: self.selecionar_arquivo(self.planilhaEntry),
        )

        planilhaBtn.grid(row=0, column=2, sticky="w", padx=5)

        dataInicialLabel = Label(forms, text="Data Inicial:", font=self.fontePadrao)
        dataInicialLabel.grid(row=1, column=0, sticky="e", pady=5, padx=5)

        self.dataInicialEntry = DateEntry(
            forms,
            width=12,
            font=self.fontePadrao,
            locale="pt_br",
            date_pattern="dd/mm/yyyy",
        )
        self.dataInicialEntry.grid(row=1, column=1, sticky="w", pady=5)

        dataFinalLabel = Label(forms, text="Data Final:", font=self.fontePadrao)
        dataFinalLabel.grid(row=2, column=0, sticky="e", pady=5, padx=5)

        self.dataFinalEntry = DateEntry(
            forms,
            width=12,
            font=self.fontePadrao,
            locale="pt_br",
            date_pattern="dd/mm/yyyy",
        )
        self.dataFinalEntry.grid(row=2, column=1, sticky="w", pady=5)

        self.executarBtn = Button(
            container,
            text="Executar",
            font=self.fontePadraoBold,
            command=lambda: self.executar_thread(self.executar),
            width=20,
        )
        self.executarBtn.grid(row=2, column=0, pady=15)

        self.pararBtn = Button(
            container,
            text="Parar",
            bg="red",
            fg="white",
            font=self.fontePadraoBold,
            command=self.forcar_parada,
            width=20,
        )

        self.pararBtn.grid(row=4, column=0, pady=5)

        progress_frame = Frame(container)
        progress_frame.grid(row=5, column=0, sticky="ew", pady=(5, 10))

        self.progressLabel = Label(progress_frame, text="", font=self.fontePadrao)
        self.progressLabel.pack(anchor="w")

        self.progressBar = ttk.Progressbar(
            progress_frame, orient="horizontal", mode="determinate", maximum=100
        )

        self.progressBar.pack(fill="x", expand=True)

        console_frame = Frame(container)
        console_frame.grid(row=6, column=0, sticky="nsew")
        container.grid_rowconfigure(6, weight=1)

        self.console = scrolledtext.ScrolledText(
            console_frame, state="normal", height=15, bg="black", fg="white"
        )

        self.console.pack(fill="both", expand=True)

        sys.stdout = TextRedirector(self.console)
        sys.stderr = TextRedirector(self.console)

        self.centralizar_janela(master)

    def centralizar_janela(self, janela):

        janela.update()

        largura = janela.winfo_width()
        altura = janela.winfo_height()

        largura_tela = janela.winfo_screenwidth()
        altura_tela = janela.winfo_screenheight()

        x = (largura_tela // 2) - (largura // 2)
        y = (altura_tela // 2) - (altura // 2)

        janela.geometry(f"+{x}+{y}")

    def selecionar_arquivo(self, entry, types=None):

        if types is None:
            types = [("Arquivos Excel", "*.xls *.xlsx"), ("Todos os arquivos", "*.*")]

        caminho = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=types,
        )

        if caminho:
            entry.delete(0, END)
            entry.insert(0, caminho)

    def executar_thread(self, comando):
        threading.Thread(target=comando, daemon=True).start()

    def forcar_parada(self):

        if not self.executando:
            return

        print("\n========== PARADA DE EMERGÊNCIA ==========\n")

        self.stop_event.set()

        try:

            py.mouseUp(button="left")
            py.mouseUp(button="right")

            py.keyUp("shift")
            py.keyUp("ctrl")
            py.keyUp("alt")

            py.moveTo(0, 0)

        except Exception:
            pass

        self.master.after(
            0,
            lambda: messagebox.showwarning(
                "Execução Interrompida", "Automação interrompida com segurança."
            ),
        )

    def validar_entrada(self):

        arquivo = self.planilhaEntry.get()
        dataInicial = self.dataInicialEntry.get()
        dataFinal = self.dataFinalEntry.get()

        if not arquivo:
            self.mostrar_erro("Erro", "Por favor, selecione uma planilha!")
            return False

        if not dataInicial:
            self.mostrar_erro("Erro", "Por favor, insira a data de início!")
            return False

        if not dataFinal:
            self.mostrar_erro("Erro", "Por favor, insira a data de fim!")
            return False
        return True

    def safe_click(self, x, y, espera=0.2):

        self.check_stop()

        largura, altura = py.size()

        if not (0 <= x <= largura and 0 <= y <= altura):
            raise ValueError(f"Coordenada inválida: ({x}, {y})")

        py.moveTo(x, y, duration=0.1)

        self.check_stop()

        py.click()

        self.esperar(espera)

    def safe_write(self, texto, espera=0.2):

        self.check_stop()

        py.write(str(texto), interval=0.01)

        self.esperar(espera)

    def safe_press(self, tecla, espera=0.2):

        self.check_stop()

        py.press(tecla)

        self.esperar(espera)

    def check_stop(self):

        if self.stop_event.is_set():
            raise InterruptedError("Execução interrompida pelo usuário")

    def esperar(self, segundos):

        intervalo = 0.05

        fim = time.time() + segundos

        while time.time() < fim:

            self.check_stop()

            time.sleep(intervalo)

    def perguntar(self, func, titulo, mensagem):

        resultado = threading.Event()
        resposta = {"valor": False}

        def mostrar():
            resposta["valor"] = func(titulo, mensagem)
            resultado.set()

        self.master.after(0, mostrar)

        resultado.wait()

        return resposta["valor"]

    def mostrar_erro(self, titulo, mensagem):

        self.master.after(0, lambda: messagebox.showerror(titulo, mensagem))

    def atualizar_progresso(self, atual, total):

        progresso = (atual / total) * 100

        self.progressBar.config(value=progresso)

        self.progressLabel.config(text=f"{atual}/{total}")

    def abrir_menu(self):

        self.safe_click(110, 37)
        self.safe_click(110, 63)
        self.safe_click(343, 129)
        self.safe_click(515, 129)

    def selecionar_funcionario(self, codigo):

        self.safe_click(428, 211)
        self.safe_click(524, 217)
        self.safe_write(str(int(codigo)))
        self.safe_press("enter")

    def novo(self):

        self.safe_click(410, 238)

    def evento(self, cod):

        self.safe_click(400, 290)
        self.safe_write(str(int(cod)))
        self.safe_press("enter")

    def parametro(self, valor):

        self.safe_click(400, 330)
        self.safe_write(str(valor).replace(",", "."))

    def datas(self, mes_atual=None, mes_inicio=None, mes_fim=None):

        self.safe_click(615, 330)

        # colocar a data inicial
        if mes_atual == mes_inicio:
            self.safe_press("1")
        elif mes_atual > mes_inicio:
            for _ in range(mes_atual - mes_inicio):
                self.safe_press("down")
        else:
            for _ in range(mes_inicio - mes_atual):
                self.safe_press("up")

        self.safe_press("tab")

        # colocar a data final
        if mes_atual == mes_fim:
            self.safe_press("1")
        elif mes_atual > mes_fim:
            for _ in range(mes_atual - mes_fim):
                self.safe_press("down")
        else:
            for _ in range(mes_fim - mes_atual):
                self.safe_press("up")

    def gravar(self):

        self.safe_click(840, 330)

    def config(self):

        janelaConfig = Toplevel(self.master)
        janelaConfig.title("Configurações")

        janelaConfig.transient(self.master)
        janelaConfig.grab_set()

        janelaConfig.minsize(450, 550)

        janelaConfig.configure(padx=20, pady=20)

        janelaConfig.grid_columnconfigure(0, weight=1)
        janelaConfig.grid_columnconfigure(1, weight=1)

        Label(janelaConfig, text="Configurações:", font=self.fontePadraoBold).grid(
            row=0, column=0, columnspan=3, pady=20
        )

        Label(
            janelaConfig,
            text="Importar configurações:",
            font=self.fontePadrao,
        ).grid(row=1, column=0, columnspan=3)

        importarEntry = Entry(janelaConfig, font=self.fontePadrao, width=30)
        importarEntry.grid(row=2, column=0, sticky="e", pady=5)

        importarBtnFile = Button(
            janelaConfig,
            text="📁",
            command=lambda: self.selecionar_arquivo(
                importarEntry,
                types=[("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")],
            ),
        )
        importarBtnFile.grid(row=2, column=1, sticky="w", padx=5)

        importarLabel = Label(
            janelaConfig, text="Configuração selecionada: Padrão", font=self.fontePadrao
        )
        importarLabel.grid(row=3, column=0, columnspan=3)

        importarBtn = Button(
            janelaConfig, text="Importar Configuração", font=self.fontePadrao
        )
        importarBtn.grid(row=4, column=0, columnspan=3, pady=5)

        Label(janelaConfig, text="Personalização:", font=self.fontePadraoBold).grid(
            row=5, column=0, columnspan=3, pady=20
        )

        personalizarBtn = Button(
            janelaConfig, text="Personalizar Configuração", font=self.fontePadrao
        )
        personalizarBtn.grid(row=6, column=0, columnspan=3, pady=5)

        Label(
            janelaConfig,
            text="Utilize a opção de personalizar configurações para ajustar os cliques de acordo com a resolução do seu sistema.",
            font=self.fontePadrao,
            wraplength=350,
            justify="left",
            anchor="w",
        ).grid(row=7, column=0, columnspan=3, pady=5)

        gerarPlanilhaBtn = Button(
            janelaConfig, text="Gerar Planilha Exemplo", font=self.fontePadrao
        )
        gerarPlanilhaBtn.grid(row=8, column=0, columnspan=3, pady=5)

        Label(
            janelaConfig,
            text="Gere uma planilha exemplo para entender melhor o formato necessário para o funcionamento do sistema.",
            font=self.fontePadrao,
            wraplength=350,
            justify="left",
            anchor="w",
        ).grid(row=9, column=0, columnspan=3, pady=5)

        frameBtnConfig = Frame(janelaConfig)
        frameBtnConfig.grid(row=10, column=0, columnspan=3, pady=25)

        Button(
            frameBtnConfig, text="Sair", command=janelaConfig.destroy, width=12
        ).pack(side=LEFT, padx=5)
        Button(
            frameBtnConfig, text="Confirmar", command=janelaConfig.destroy, width=12
        ).pack(side=LEFT, padx=5)

        self.centralizar_janela(janelaConfig)

    def executar(self):

        if self.executando:
            return

        if not self.validar_entrada():
            return

        self.stop_event.clear()

        self.executando = True
        self.master.after(0, lambda: self.executarBtn.config(state="disabled"))

        try:

            py.FAILSAFE = True
            py.PAUSE = 0.03

            # automação

            CAMINHO = self.planilhaEntry.get()
            DATA_INICIO = self.dataInicialEntry.get_date()
            DATA_FIM = self.dataFinalEntry.get_date()

            mes_atual = datetime.now().month
            mes_inicio = DATA_INICIO.month
            mes_fim = DATA_FIM.month

            cod_atual = None
            nome_atual = None

            print(f"Executando com: {CAMINHO}, {DATA_INICIO}, {DATA_FIM}\n")

            df = pd.read_excel(CAMINHO, header=3)

            df = df.rename(
                columns={
                    "Código": "codigo",
                    "Nome": "nome",
                    "Evento": "evento",
                    "Ref./Valor": "valor",
                }
            )

            colunas_necessarias = ["codigo", "nome", "evento", "valor"]

            for col in colunas_necessarias:
                if col not in df.columns:
                    raise ValueError(f"Coluna obrigatória ausente: {col}")

            df = df[df["nome"].notna()]
            df = df[df["valor"].notna()]

            df = df.sort_values(by=["codigo"])

            total = len(df)

            if total == 0:
                self.perguntar(
                    messagebox.showwarning,
                    "Aviso",
                    "Nenhum dado encontrado na planilha.",
                )
                return

            print("Aguardando usuário abrir o sistema...")

            confirmado = self.perguntar(
                messagebox.askokcancel,
                "Confirmação",
                "Abra o sistema e clique em OK para iniciar.\nVocê terá 5 segundos após clicar em OK para deixar o sistema aberto.",
            )

            if not confirmado:

                print("Usuário cancelou.")

                return

            self.esperar(5)

            self.abrir_menu()

            for i, row in enumerate(df.itertuples(index=False), start=1):

                self.master.after(
                    0, lambda i=i, total=total: self.atualizar_progresso(i, total)
                )

                codigo = row.codigo
                nome = row.nome
                evento_cod = row.evento
                valor = row.valor

                print(f"Processando: {codigo} - {nome} - {evento_cod} - {valor}")

                if cod_atual != codigo:

                    if cod_atual is not None:

                        continuar_execucao = self.perguntar(
                            messagebox.askyesno,
                            "Confirmação",
                            f"Deseja continuar?\nOs lançamentos do funcionário ({cod_atual}) {nome_atual} foram finalizados.\nAproveite para corrigir os lançamentos se necessário.",
                        )

                        if not continuar_execucao:

                            print("Usuário cancelou.")

                            return

                        else:

                            print("Continuando execução...")

                    self.selecionar_funcionario(codigo)

                    cod_atual = codigo
                    nome_atual = nome

                self.novo()
                self.evento(evento_cod)
                self.parametro(valor)
                self.datas(mes_atual, mes_inicio, mes_fim)
                self.gravar()

        except InterruptedError:

            print("Execução interrompida.")

        except py.FailSafeException:

            print("FAILSAFE acionado.")

        except Exception:
            import traceback

            traceback.print_exc()

        finally:

            self.executando = False

            self.stop_event.clear()

            self.master.after(0, lambda: self.executarBtn.config(state="normal"))


root = Tk()
root.title("Lançamento de Apontamentos")
Application(root)
root.mainloop()
