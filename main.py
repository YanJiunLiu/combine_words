import tkinter as tk
from tkinter import ttk, filedialog
import tkinterDnD  # Importing the tkinterDnD module
from functools import partial
from combine import combined


def drop(event):
    # This function is called, when stuff is dropped into a widget
    if len(labels):
        for index, l in enumerate(labels):
            l.destroy()
            up_buttons[index].destroy()
            del_buttons[index].destroy()
        labels.clear()
        up_buttons.clear()
        del_buttons.clear()
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
                          textvar=StringVars[i], padding=10, relief="solid")
        label.configure(anchor="center", font=('Arial bold', 12))
        labels.append(label)
        up_button = ttk.Button(
            scrollable_frame,
            text='up',
            command=partial(move_up, i), width=5
        )
        up_buttons.append(up_button)
        del_button = ttk.Button(
            scrollable_frame,
            text='del',
            command=partial(delete, i), width=5
        )
        del_buttons.append(del_button)
    for index, l in enumerate(labels):
        l.pack(fill="both", expand=True, padx=10, pady=10)
        del_buttons[index].pack(expand=True)
        up_buttons[index].pack(expand=True)

    scrollbar.pack(side="right", fill="y")
    canvas2.pack(fill="both", expand=True, padx=10, pady=10)
    scrollbar_x.pack(side="bottom", fill="both")


def delete():
    pass


def move_up(i):
    tmp = StringVars[i]
    StringVars[i] = StringVars[i - 1]
    StringVars[i - 1] = tmp
    # refresh
    drop(None)


def save_file_as():
    # If there is no file path specified, prompt the user with a dialog which
    # allows him/her to select where they want to save the file
    filename = filedialog.asksaveasfilename(
        title='Save file',
        filetypes=(
            ('text files', '*.docx'),
            ('text files', '*.doc')
        )
    )

    save_label['text'] = f"{filename}"


def combine():
    if len(StringVars) < 2:
        combine_label['text'] = f"Please give more than one docxes."
        return
    stringvars = [StringVar.get() for StringVar in StringVars]
    if save_label['text'] and stringvars:
        combined(files=stringvars, dst=save_label['text'])
        combine_label['text'] = f"Success."
        return
    else:
        combine_label['text'] = f"Please give the saving path."
        return


root = tkinterDnD.Tk()
root.title("docx_combiner")
# root.maxsize(600, 600)
StringVars = []
labels = []
up_buttons = []
del_buttons = []
# frame1 block
frame1 = tk.Frame(root)
frame1.pack(fill="both", expand=True)
canvas1 = tk.Canvas(frame1, bg="white")
canvas1.pack(fill="both", expand=True)

## frame1 block -- drop block
stringvar = tk.StringVar()
stringvar.set('Drop here!')
drop_label = ttk.Label(canvas1, ondrop=drop, textvar=stringvar, padding=40, relief="solid")
drop_label.configure(anchor="center", font=('Helvatical bold', 20))
drop_label.pack(fill="both", expand=True, padx=10, pady=10)

## frame1 block -- save block
save_label = ttk.Label(canvas1, padding=1, relief="solid")
save_label.configure(anchor="center", font=('Helvatical bold', 12))
save_label.pack(fill="both", expand=True, padx=10, pady=10)
save_button = ttk.Button(
    canvas1,
    text='save as',
    command=partial(save_file_as)
)
save_button.pack()

## frame1 block -- combine block
combine_label = ttk.Label(canvas1, padding=0.5, relief="solid")
combine_label.pack(fill="both", expand=True, padx=10, pady=10)
combine_button = ttk.Button(
    canvas1,
    text='confirm to combine',
    command=partial(combine)
)
combine_button.pack()
######################################################################################################################
interval_label = tk.Label(canvas1, height=1, bg="white")
interval_label.pack(fill="both")
######################################################################################################################
# frame2 block
frame2 = tk.Frame(root)

canvas2 = tk.Canvas(frame2, bg="black")
scrollbar = ttk.Scrollbar(frame2, orient="vertical", command=canvas2.yview)
scrollbar_x = ttk.Scrollbar(frame2, orient="horizontal", command=canvas2.xview)

scrollable_frame = ttk.Frame(canvas2)
scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas2.configure(
        scrollregion=canvas2.bbox("all")
    )
)

canvas2.create_window((0, 0), window=scrollable_frame, anchor="nw", width=2000)
canvas2.configure(yscrollcommand=scrollbar.set, xscrollcommand=scrollbar_x.set)
frame2.pack(fill="both", expand=True)
root.mainloop()
