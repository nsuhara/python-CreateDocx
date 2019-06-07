import os
import tkinter as tk
from tkinter import filedialog as fdialog
from tkinter import messagebox as mdialog

from model import Docx


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.create_widgets()

    def set_title(self):
        self.master.title('Create Docx')

    def set_menu_bar(self):
        self.menu_bar = tk.Menu(self.master)
        self.master.config(menu=self.menu_bar)
        file_menu = tk.Menu(self.menu_bar)
        file_menu.add_command(label='Exit', command=self.master.quit)
        self.menu_bar.add_cascade(label='File', menu=file_menu)

    def select_file(self, entry):
        entry.delete(0, tk.END)
        entry.insert(0, fdialog.askopenfilename(initialdir=os.getcwd()))

    def create_docx(self, json_url, template_url):
        if not os.path.exists(json_url) or not os.path.exists(template_url):
            mdialog.showerror('Error', 'Please select JSON and Template.')
            return

        docx = Docx(json_url=json_url, template_url=template_url)
        docx.render()

    def set_body(self):
        tk.Label(self.master, text='JSON:').grid(row=0, column=0)
        entry_json = tk.Entry(self.master)
        entry_json.grid(row=0, column=1, pady=5)
        tk.Button(self.master, text='Select...',
                  command=lambda: self.select_file(entry_json)).grid(row=0, column=2)

        tk.Label(self.master, text='Template:').grid(row=1, column=0)
        entry_template = tk.Entry(self.master)
        entry_template.grid(row=1, column=1, pady=5)
        tk.Button(self.master, text='Select...',
                  command=lambda: self.select_file(entry_template)).grid(row=1, column=2)

        tk.Button(self.master, text='Create', width=30,
                  command=lambda: self.create_docx(entry_json.get(), entry_template.get())).grid(row=2, column=0, columnspan=3)

    def create_widgets(self):
        self.master.geometry()
        self.entry = tk.Entry(self.master)

        self.set_title()
        self.set_menu_bar()
        self.set_body()


# fix tkinter bug start
def fix_bug():
    width_height = root.winfo_geometry().split('+')[0].split('x')
    width = int(width_height[0])
    height = int(width_height[1])
    root.geometry('{}x{}'.format(width+1, height+1))
# fix tkinter bug end


if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    # fix tkinter bug start
    root.update()
    root.after(0, fix_bug)
    # fix tkinter bug end
    app.mainloop()
