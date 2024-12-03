import os
from tkinter import Tk, Button, Label, filedialog
from tkinter.messagebox import showinfo, showerror
from tkinter import font
from PIL import Image
from fpdf import FPDF
import docx2txt
from PyPDF2 import PdfMerger

# Função para converter imagens em PDF
def image_to_pdf(image_path, output_pdf):
    image = Image.open(image_path)
    if image.mode != 'RGB':
        image = image.convert('RGB')
    image.save(output_pdf, 'PDF')

# Função para converter documentos Word (.docx) para PDF
def docx_to_pdf(docx_path, output_pdf):
    try:
        text = docx2txt.process(docx_path)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, text)
        pdf.output(output_pdf)
    except Exception as e:
        showerror("Erro", f"Erro ao converter o documento Word: {e}")

# Função para converter arquivos PDF em PDF
def pdf_to_pdf(pdf_path, output_pdf):
    try:
        merger = PdfMerger()
        merger.append(pdf_path)
        merger.write(output_pdf)
        merger.close()
    except Exception as e:
        showerror("Erro", f"Erro ao copiar o PDF: {e}")

# Função principal para detectar o tipo de arquivo e chamar a função apropriada
def convert_to_pdf(input_file, output_pdf):
    file_extension = input_file.lower().split('.')[-1]

    if file_extension in ['jpg', 'jpeg', 'png', 'bmp', 'gif']:
        image_to_pdf(input_file, output_pdf)
    elif file_extension == 'docx':
        docx_to_pdf(input_file, output_pdf)
    elif file_extension == 'pdf':
        pdf_to_pdf(input_file, output_pdf)
    else:
        showerror("Erro", "Tipo de arquivo não suportado. Apenas imagens, documentos Word e PDFs são aceitos.")

# Função para abrir o seletor de arquivos
def open_file():
    global input_file
    input_file = filedialog.askopenfilename(title="Selecione o arquivo", filetypes=[("Todos os arquivos", "*.*")])
    if input_file:
        file_label.config(text=f"Arquivo selecionado: {os.path.basename(input_file)}")

# Função para converter o arquivo selecionado
def convert_file():
    if input_file:
        # Pergunta onde salvar o PDF
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_pdf:
            try:
                convert_to_pdf(input_file, output_pdf)
                showinfo("Sucesso", f"Arquivo convertido para PDF:\n{output_pdf}")
            except Exception as e:
                showerror("Erro", f"Ocorreu um erro durante a conversão: {e}")
    else:
        showerror("Erro", "Por favor, selecione um arquivo primeiro.")

# Configuração da interface gráfica
root = Tk()
root.title("Conversor de Arquivos para PDF")
root.geometry("500x300")
root.config(bg="#f7f7f7")  # Cor de fundo suave

input_file = ""

# Fontes personalizadas para a interface
font_label = font.Font(family="Arial", size=12)
font_button = font.Font(family="Arial", size=12, weight="bold")

# Criando a interface
file_label = Label(root, text="Nenhum arquivo selecionado", wraplength=300, bg="#f7f7f7", font=font_label, fg="#333")
file_label.pack(pady=20)

# Botão de seleção de arquivo
select_button = Button(root, text="Selecionar Arquivo", command=open_file, bg="#4CAF50", fg="white", font=font_button, relief="solid", borderwidth=1, padx=10, pady=5)
select_button.pack(pady=10)

# Botão de conversão
convert_button = Button(root, text="Converter", command=convert_file, bg="#008CBA", fg="white", font=font_button, relief="solid", borderwidth=1, padx=10, pady=5)
convert_button.pack(pady=20)

# Iniciar a interface gráfica
root.mainloop()
