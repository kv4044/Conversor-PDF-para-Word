import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as font
import os
from pdf2docx import Converter

root = tk.Tk()
root.title("Pdf To Docx")


def add_file():
    file_path = filedialog.askopenfilename(initialdir="/", title="Select file")
    convert_file(file_path)


def convert_file(file_path):
    if not file_path:
        messagebox.showerror("Error", "Please select a PDF file")
        return
    save_location = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    if not save_location:
        return
    success_text.config(text="Converting... wait")
    try:
        cv = Converter(file_path)
        cv.convert(save_location, start=0, end=None)
        cv.close()
        success_text.config(text="Pdf has been successfully converted")
        success_text.pack()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


canvas = tk.Canvas(root, height=200, width=400)
canvas.pack()

frame = tk.Frame(root)
frame.place(relwidth=.8, relheight=.8, relx=.1, rely=.1)

label_titulo = tk.Label(frame, text="PDF to Word converter.", font=("Arial", 14))
label_titulo.pack(pady=10)

button = tk.Button(frame, text="OPEN FILE ", padx=10, pady=5, bg="#ff2200", fg="#fff", command=add_file)

success_text = tk.Label(frame, text="Pdf has been successfully converted", font=16)

button_font = font.Font(size=15)
button["font"] = button_font
button.pack()

root.mainloop()
