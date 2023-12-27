import time
import tkinter as tk
from tkinter import ttk, scrolledtext

class GUI:
    @staticmethod
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Pesquisa de Planilhas")

        self.search_label = ttk.Label(self.root, text="Digite o valor a ser procurado nas planilhas:")
        self.search_label.pack(pady=5)

        self.search_entry = ttk.Entry(self.root)
        self.search_entry.pack(pady=5)

        self.month_label = ttk.Label(self.root, text="Digite o mês a ser procurado (por exemplo, DEZEMBRO):")
        self.month_label.pack(pady=5)

        self.month_entry = ttk.Entry(self.root)
        self.month_entry.pack(pady=5)

        self.result_text = scrolledtext.ScrolledText(self.root, width=40, height=10)
        self.result_text.pack(pady=10)

        self.status_label = ttk.Label(self.root, text="Status da Aplicação:")
        self.status_label.pack(pady=5)

        self.status_var = tk.StringVar()
        self.status_var.set("Aguardando Início")
        self.status_display = ttk.Label(self.root, textvariable=self.status_var)
        self.status_display.pack(pady=5)

        self.retry_delay_seconds = 5
        self.max_wait_time_seconds = 180
        self.requests_count = 0
        self.start_time = time.time()

        # Inicia a interface gráfica
        self.create_gui()

    def create_gui(self):
        # Criação da interface gráfica
        search_button = ttk.Button(self.root, text="Iniciar Pesquisa", command=self.start_search)
        search_button.pack(pady=10)

        stop_button = ttk.Button(self.root, text="Encerrar", command=self.root.destroy)
        stop_button.pack(pady=10)

        # Inicia a interface gráfica
        self.root.mainloop()
