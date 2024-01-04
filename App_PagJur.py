import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
import re
import os
import shutil

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PagJur")

        self.label = tk.Label(root, text="Selecione a pasta com os arquivos PDF:")
        self.label.pack(pady=10)

        self.folder_path = tk.StringVar()
        self.entry = tk.Entry(root, textvariable=self.folder_path, width=40)
        self.entry.pack(pady=10)

        self.browse_button = tk.Button(root, text="Procurar", command=self.browse_folder)
        self.browse_button.pack(pady=10)

        self.process_button = tk.Button(root, text="Unificar PDFs", command=self.process_pdfs)
        self.process_button.pack(pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def extract_and_move_cover(self, file_path):
        try:
            pdf_document = fitz.open(file_path)
            first_page = pdf_document[0]
            text = first_page.get_text()
            pdf_document.close()

            if 'Pagamentos para GEAFI - Requisições' in text:
                # Encontra todas as sequências de 6 dígitos no texto
                all_matches = re.findall(r'\b\d{6}\b', text)
                
                for folder_number in all_matches:
                    if folder_number != '109106' and folder_number != '103812':
                        source_folder = os.path.dirname(file_path)
                        destination_folder = os.path.join(source_folder, folder_number)

                        if not os.path.exists(destination_folder):
                            os.makedirs(destination_folder)
                            print(f"Pasta criada: {destination_folder}")

                        destination_path = os.path.join(destination_folder, os.path.basename(file_path))
                        shutil.move(file_path, destination_path)
                        print(f"Arquivo de capa movido para {destination_path}")
                        break  # Para após a primeira correspondência

        except Exception as e:
            messagebox.showerror("Erro na Extração e Movimentação", f"Erro ao processar o arquivo {file_path}: {str(e)}")

    def move_covers(self):
        folder_path = self.folder_path.get()

        # Executa a extração da capa e movimento para a pasta correspondente
        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
        for pdf_file in pdf_files:
            file_path = os.path.join(folder_path, pdf_file)
            if os.path.isfile(file_path):
                self.extract_and_move_cover(file_path)

    def merge_pdfs_recursive(self, base_folder):
        try:
            folder_files = {}

            for folder_name, subfolders, files in os.walk(base_folder):
                for file in files:
                    file_path = os.path.join(folder_name, file)

                    if file.lower().endswith('.pdf'):
                        if folder_name not in folder_files:
                            folder_files[folder_name] = [file_path]
                        else:
                            folder_files[folder_name].append(file_path)

            for folder_name, file_paths in folder_files.items():
                if folder_name != '109106':
                    merged_pdf = PdfWriter()
                    capa_file = next((f for f in file_paths if 'Pagamentos para GEAFI - Requisições' in f), None)

                    if capa_file:
                        with open(capa_file, 'rb') as capa:
                            pdf_reader = PdfReader(capa)
                            for page_num in range(len(pdf_reader.pages)):
                                page = pdf_reader.pages[page_num]
                                merged_pdf.add_page(page)

                    for file_path in file_paths:
                        if file_path != capa_file:
                            with open(file_path, 'rb') as file:
                                pdf_reader = PdfReader(file)
                                for page_num in range(len(pdf_reader.pages)):
                                    page = pdf_reader.pages[page_num]
                                    merged_pdf.add_page(page)

                    output_path = os.path.join(base_folder, f'{folder_name}.pdf')

                    with open(output_path, 'wb') as output_file:
                        merged_pdf.write(output_file)

                    print(f"PDF unificado em {output_path}")

        except Exception as e:
            messagebox.showerror("Erro na Unificação de PDFs", f"Erro ao unificar PDFs na pasta {base_folder}: {str(e)}")

    def process_pdfs(self):
        self.move_covers()

        folder_path = self.folder_path.get()

        # Ordena as pastas antes de unificar os PDFs
        subfolders = sorted([f.path for f in os.scandir(folder_path) if f.is_dir()])

        # Executa a unificação dos PDFs
        for subfolder in subfolders:
            self.merge_pdfs_recursive(subfolder)

        # Exibe a mensagem de conclusão
        messagebox.showinfo("Concluído", "Unificação foi realizada, obrigada por hoje!")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()
