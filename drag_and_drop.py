import tkinter as tk
from tkinter import ttk
import tkinterDnD  # Importing the tkinterDnD module
from functools import partial

root = tkinterDnD.Tk()
root.title("docx_combiner")
root.maxsize(600, 600)


def drop(event):
    # This function is called, when stuff is dropped into a widget
    if len(labels):
        for index, l in enumerate(labels):
            l.destroy()
            up_buttons[index].destroy()
        labels.clear()
        up_buttons.clear()
    if event:
        files = event.data.split(' ')
        start = len(StringVars)
        for _ in files:
            StringVars.append(tk.StringVar())
    for i, StringVar in enumerate(StringVars):
        if event:
            if start <= i:
                StringVar.set(files[i - start])
        label = ttk.Label(scrollable_frame,
                          textvar=StringVars[i], padding=15, relief="solid")
        labels.append(label)
        up_button = ttk.Button(
            scrollable_frame,
            text='up',
            command=partial(move_up, i), width=5
        )
        up_buttons.append(up_button)
    for index, l in enumerate(labels):
        l.pack(fill="both", expand=True, padx=10, pady=10)
        up_buttons[index].pack(expand=True)

    frame.pack()
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")


def move_up(i):
    tmp = StringVars[i]
    StringVars[i] = StringVars[i-1]
    StringVars[i-1] = tmp
    # refresh
    drop(None)

def combine():
    pass


# Without DnD hook you need to register the widget for every purpose,
# and bind it to the function you want to call
stringvar = tk.StringVar()
stringvar.set('Drop here or drag from here!')
StringVars = []
labels = []
up_buttons = []
frame = ttk.Frame(root)
canvas = tk.Canvas(frame)
scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

canvas.configure(yscrollcommand=scrollbar.set)
label = ttk.Label(root, ondrop=drop,  textvar=stringvar, padding=50, relief="solid")
label.pack(fill="both", expand=True, padx=10, pady=10)
open_button = ttk.Button(
    root,
    text='combine',
    command=combine()
)
open_button.pack(expand=True)

root.mainloop()
