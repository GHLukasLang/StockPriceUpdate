import tkinter as tk
from tkinter import filedialog

from functions import handle_file, back_up

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.frames = {}
        container = tk.Frame(self)
        container.pack(fill="both", expand=True)
        self.title("StockPriceUpdater")



        for F in (IntroPage, FileSelectPage, ResultPage):
            frame = F(container, self)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("IntroPage")

    def show_frame(self, page_name):
        self.frames[page_name].tkraise()

class IntroPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.title = "StockPriceUpdater"
        tk.Label(self, text="Welcome to the StockPriceUpdater!\n\n\nPlease proceed to pick your .xlsx file.\n\nAfter making a backup-copy, it will scan the first sheet for stock names,\nfetch their current prices and add those to a new sheet.").pack(pady=10)
        tk.Button(self, text="Next", command=lambda: controller.show_frame("FileSelectPage")).pack()

class FileSelectPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        tk.Label(self, text="Please pick a file:").pack(pady=10)
        tk.Button(self, text="Browse", command=self.pick_file).pack()
        self.controller = controller

    def pick_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            back_up(file_path)
            handle_file(file_path)
            self.controller.show_frame("ResultPage")

class ResultPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        tk.Label(self, text="File processing result here!").pack(pady=10)
        tk.Button(self, text="Finish", command=self.quit).pack()

if __name__ == "__main__":
    app = App()
    app.mainloop()