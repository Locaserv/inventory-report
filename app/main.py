import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os

from generate_xlsx import PDFToExcelConverter

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class EstoqueApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Relatório de Estoque LOCASERV")
        self.geometry("500x300")

        self.selected_files = []

        self.label = ctk.CTkLabel(self, text="Selecione os arquivos PDF:")
        self.label.pack(pady=10)

        self.select_button = ctk.CTkButton(self, text="Selecionar PDFs", command=self.select_pdfs)
        self.select_button.pack()

        self.generate_button = ctk.CTkButton(self, text="Gerar Relatório", command=self.start_report_thread, state="disabled")
        self.generate_button.pack(pady=10)

        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.pack(pady=5)

    def select_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if files:
            self.selected_files = list(files)
            file_names = "\n".join([os.path.basename(file) for file in self.selected_files])
            self.status_label.configure(
                text=f"Arquivos selecionados:\n{file_names}"
            )
            self.generate_button.configure(state="normal")

    def start_report_thread(self):
        if not self.selected_files:
            messagebox.showwarning("Aviso", "Nenhum PDF selecionado.")
            return
        self.status_label.configure(text="⏳ Gerando relatório, por favor aguarde...")
        threading.Thread(target=self.generate_report).start()

    def generate_report(self):
        try:
            converter = PDFToExcelConverter()
            for pdf in self.selected_files:
                converter.add_pdf_path(pdf)
            converter.generate_report()
            caminho = converter.get_xlsx_path()

            self.status_label.configure(text="✅ Relatório gerado com sucesso!")
            os.startfile(caminho)  # Abre o arquivo Excel ao final
        except Exception as e:
            self.status_label.configure(text="❌ Erro ao gerar relatório.")
            messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    app = EstoqueApp()
    app.mainloop()