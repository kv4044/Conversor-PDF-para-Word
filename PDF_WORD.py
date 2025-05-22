import customtkinter as ctk
from customtkinter import CTkTabview
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx2pdf import convert
from reportlab.pdfgen import canvas
from PIL import Image

# Configuração inicial
ctk.set_appearance_mode("Dark")  # ou "Light", "System"
ctk.set_default_color_theme("green")  # cores: blue, green, dark-blue

PDF_DOCX = "PDF  →  DOCX"
DOCX_PDF = "DOCX  →  PDF"
TXT_PDF = "TXT  →  PDF"
JPG_PDF = "JPG  →  PDF"

label_de_sucesso = {}

app = ctk.CTk()
app.title("Conversor de arquivos")
app.geometry("500x300")

tabview = CTkTabview(master=app)
tabview.pack(pady=20, fill="both", expand=True)

tabview.add(PDF_DOCX)
tabview.add(DOCX_PDF)
tabview.add(TXT_PDF)
tabview.add(JPG_PDF)


# Função para converter
def converter_pdf_para_docx():
    file_path = filedialog.askopenfilename(title="Selecione um arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    if file_path:
        save_location = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if not save_location:
            return
        label_de_sucesso[PDF_DOCX].configure(text="Convertendo... Aguarde")
        try:
            cv = Converter(file_path)
            cv.convert(save_location, start=0, end=None)
            cv.close()
            label_de_sucesso[PDF_DOCX].configure(text="Conversão concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            label_de_sucesso[PDF_DOCX].configure(text="")


def converter_docx_para_pdf():
    file_path = filedialog.askopenfilename(title="Selecione um arquivo DOCX", filetypes=[("Arquivos docx", "*.docx")])
    if file_path:
        save_location = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Arquivos PDF", "*.pdf")])
        if not save_location:
            return
        label_de_sucesso[DOCX_PDF].configure(text="Convertendo... Aguarde")
        try:
            convert(file_path, save_location)
            label_de_sucesso[DOCX_PDF].configure(text="Conversão concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            label_de_sucesso[DOCX_PDF].configure(text="")


def converter_txt_para_pdf():
    file_path = filedialog.askopenfilename(title="Selecione um arquivo TXT", filetypes=[("Arquivo txt", "*.txt")])
    if file_path:
        save_location = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("ARquivo PDF", "*.pdf")])
        if not save_location:
            return
        label_de_sucesso[TXT_PDF].configure(text="Convertendo... Aguarde!")
        try:
            c = canvas.Canvas(save_location)
            x = 40
            y = 800

            with open(file_path, "r", encoding="utf-8") as f:
                for linha in f:
                    c.drawString(x, y, linha.strip())
                    y -= 15
                    if y < 40:
                        c.showPage()
                        y = 800

            c.save()
            label_de_sucesso[TXT_PDF].configure(text="Conversão concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            label_de_sucesso[DOCX_PDF].configure(text="")


def converter_jpg_para_pdf():
    file_path = filedialog.askopenfilename(title="Selecione um arquivo JPG", filetypes=[("Arquivo JPG", "*.jpg")])
    if file_path:
        save_location = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("ARquivo PDF", "*.pdf")])
        if not save_location:
            return
        label_de_sucesso[TXT_PDF].configure(text="Convertendo... Aguarde!")
        try:
            img = Image.open(file_path)
            img_rgb = img.convert("RGB")
            img_rgb.save(save_location, "PDF")
            label_de_sucesso[JPG_PDF].configure(text="Conversão concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            label_de_sucesso[DOCX_PDF].configure(text="")


def criar_buttom_label(tab, text_button, text_label, comando):
    aba = tabview.tab(tab)
    # Título
    title_label = ctk.CTkLabel(master=aba,
                               text=f"Conversor de {text_label}", font=("Arial", 20, "bold"))
    title_label.pack(pady=30)

    # Botão de seleção
    select_button = ctk.CTkButton(
        master=aba,
        text=f"{text_button}",
        font=("Arial", 16),
        fg_color="#4CAF50",
        hover_color="#45a049",
        text_color="white",
        corner_radius=8,
        command=comando
    )
    select_button.pack(pady=10)
    # Label de sucesso específica para essa aba
    label = ctk.CTkLabel(master=aba, text="", font=("Arial", 14))
    label.pack(pady=10)

    # Guardar no dicionário para usar depois
    label_de_sucesso[tab] = label


criar_buttom_label(PDF_DOCX, "Selecionar PDF", "PDF para DOCX", converter_pdf_para_docx)
criar_buttom_label(DOCX_PDF, "Selecionar DOCX", "DOCX para PDF", converter_docx_para_pdf)
criar_buttom_label(TXT_PDF, "Selecionar TXT", "TXT para PDF", converter_txt_para_pdf)
criar_buttom_label(JPG_PDF, "Selecionar JPG", "JPG para PDF", converter_jpg_para_pdf)

# Iniciar app
app.mainloop()
