from tkinter import *
from tkinter import ttk
import tkinter as tk
from tiun import create_files
from tkinter import filedialog
import os

root = Tk()

root.title("РегионЭнергоСеть")
root.geometry("600x400")


def bb():
    tf = filedialog.askopenfilename(
        initialdir=os.getcwd(),  # insert your own path
        title="Open Text file"
    )

    if len(tf.split("/")[-1].split('.')[-1]) > 0:
        if tf.split("/")[-1].split('.')[-1] not in ('xlsx'):
            button_exel['text'] = "Ошибка: Файл должен быть формата Exel"
            button_exel.configure(background="brown")
        else:
            button_exel['text'] = f'Загружено: \n{tf.split("/")[-1]}'
            button_exel.configure(background="green")

def gg():
    tf = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Open Text file"
    )

    if len(tf.split("/")[-1].split('.')[-1]) > 0:
        if tf.split("/")[-1].split('.')[-1] not in ("doc", "docx"):
            button_word['text'] = "Ошибка: Файл должен быть формата Word"
            button_word.configure(background="brown")
        else:
            button_word['text'] = f'Загружено: \n{tf.split("/")[-1]}'
            button_word.configure(background="green")

def main():
    if button_exel.cget("bg") == button_word.cget("bg") == "green":

        try:
            button_main.configure(background="green")
            create_files(button_exel['text'].split(":")[1][2:], button_word['text'].split(":")[1][2:])
            label.configure(text="Файлы созданы")
        except Exception as e:
            button_main.configure(background="brown")
            label.configure(text=f"Ошибка при создании файлов:\n{e}")

    else:
        button_main['text'] = f'Ошибка загружаемых файлов'
        button_main.configure(background="brown")

label = Label(root, text="", font=("Arial Bold", 12), height=150)
label.pack()

button_exel = tk.Button(text='Файл ексель', command=bb, bg='grey', fg='white')
button_exel.place(relx=0.5, rely=0.1, anchor=CENTER)

button_word = tk.Button(text='Файл ворд', command=gg, bg='grey', fg='white')
button_word.place(relx=0.5, rely=0.2, anchor=CENTER)

button_main = tk.Button(text='Начать создание файлов', command=main, bg='grey', fg='white')
button_main.place(relx=0.5, rely=0.4, anchor=CENTER)

root.mainloop()


