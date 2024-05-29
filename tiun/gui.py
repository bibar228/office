from tkinter import *
from tkinter import ttk
import tkinter as tk
from tiun import create_files
from tkinter import filedialog

root = Tk()

root.title("РегионЭнергоСеть")
root.geometry("300x400")
canvas1 = tk.Canvas(root, width=500, height=400)
canvas1.pack()

def bb():
    tf = filedialog.askopenfilename(
        initialdir="C:/",  # insert your own path
        title="Open Text file"
    )
    button_exel['text'] = f'Загружено: \n{tf.split("/")[-1]}'
    print(tf.split("/")[-1])

def gg():
    tf = filedialog.askopenfilename(
        initialdir="C:/",  # insert your own path
        title="Open Text file"
    )
    button_word['text'] = f'Загружено: \n{tf.split("/")[-1]}'
    print(tf.split("/")[-1])

button_exel = tk.Button(text='Файл ексель', command=bb, bg='brown', fg='white')
canvas1.create_window(150, 150, window=button_exel)
button_exel.place(relx=0.5, rely=0.5, anchor=CENTER)

button_word = tk.Button(text='Файл ворд', command=gg, bg='brown', fg='white')
canvas1.create_window(150, 150, window=button_exel)
button_word.place(relx=0.5, rely=0.5, anchor=CENTER)

root.mainloop()

# entry = ttk.Entry()
# entry.pack(anchor=NW, padx=6, pady=6)
#
# btn_word = ttk.Button(text="Название файла ворд", command=show_message)
# btn_word.pack(anchor=NW, padx=6, pady=6)
#
# label = ttk.Label()
# label.pack(anchor=NW, padx=6, pady=6)

root.mainloop()



# tf = filedialog.askopenfilename(
#     initialdir="C:/",  # insert your own path
#     title="Open Text file"
# )
#
# print(tf.split("/")[-1])