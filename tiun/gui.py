from tkinter import *
from tiun import create_files

root = Tk()

root.title("SpeedTest")
root.geometry("300x400")

button = Button(root, text="Нажатать чтобы начать", font=40, command=create_files)
button.pack(side=BOTTOM, pady=40)

root.mainloop()