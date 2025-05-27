import os
import zipfile
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import logging
import ctypes
import sys

if sys.platform == "win32":
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

logging.getLogger("pdfminer").setLevel(logging.ERROR)

def extrair_tabela(pdf_path, status_label, progress_bar):
    status_label.config(text="Lendo páginas do PDF em memória...")
    tables = []
    start_time = time.time()

    with pdfplumber.open(pdf_path) as pdf:
        paginas = pdf.pages  
        num_pages = len(paginas)

        for i, page in enumerate(paginas):
            try:
                tabela = page.extract_table()
                if tabela:
                    tables.extend(tabela)
            except:
                pass

            elapsed = time.time() - start_time
            media = elapsed / (i + 1)
            restante = media * (num_pages - (i + 1))

            progress = int(((i + 1) / num_pages) * 100)
            progress_bar["value"] = progress
            status_label.config(text=f"Página {i + 1}/{num_pages} processada - Tempo restante: {restante:.1f}s")
            progress_bar.update()

    if not tables:
        status_label.config(text="Nenhuma tabela encontrada no PDF.")
        return None

    df = pd.DataFrame(tables)
    df.columns = df.iloc[0]
    df = df[1:]
    df.replace({"OD": "Odontologia", "AMB": "Ambulatorial"}, inplace=True)

    csv_filename = "Tabela_Extraida.csv"
    df.to_csv(csv_filename, index=False, encoding="utf-8")
    status_label.config(text=f"CSV salvo como {csv_filename}")
    return csv_filename, df.head()

def compactar_csv(csv_filename, status_label):
    zip_filename = "Tabela_Compactada.zip"
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(csv_filename)
    os.remove(csv_filename)
    status_label.config(text=f"Compactado como {zip_filename}")
    return zip_filename

def iniciar_extracao(status_label, progress_bar, preview_text):
    caminho_pdf = filedialog.askopenfilename(title="Selecione um PDF", filetypes=[("PDF Files", "*.pdf")])
    if not caminho_pdf:
        return

    def tarefa():
        try:
            status_label.config(text="Iniciando...")
            progress_bar["value"] = 0
            preview_text.delete("1.0", tk.END)

            resultado = extrair_tabela(caminho_pdf, status_label, progress_bar)
            if resultado:
                csv_file, preview_df = resultado
                compactar_csv(csv_file, status_label)
                preview_text.insert(tk.END, preview_df.to_string(index=False))
                messagebox.showinfo("Concluído", "Transformação finalizada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            status_label.config(text="Erro ao processar o PDF.")

    threading.Thread(target=tarefa).start()

janela = tk.Tk()
janela.title("Transformador de PDF → CSV (Patrick)")
janela.geometry("500x420")
janela.resizable(False, False)

tk.Label(janela, text="Selecione um PDF para extrair a tabela", font=("Arial", 14)).pack(pady=10)

tk.Button(janela, text="Escolher PDF", command=lambda: iniciar_extracao(status_label, progress_bar, preview_text), font=("Arial", 12)).pack(pady=5)

progress_bar = ttk.Progressbar(janela, length=400, mode='determinate')
progress_bar.pack(pady=5)

status_label = tk.Label(janela, text="", font=("Arial", 10), fg="blue")
status_label.pack(pady=5)

tk.Label(janela, text="Preview das primeiras linhas:", font=("Arial", 10)).pack()
preview_text = tk.Text(janela, height=10, width=60)
preview_text.pack(pady=5)

janela.mainloop()
