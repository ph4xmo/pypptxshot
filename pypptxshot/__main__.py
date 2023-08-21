import tkinter as tk

from .pypptxshot import PyPPTXShot

if __name__ == "__main__":
    app = tk.Tk()
    pptx_shot = PyPPTXShot(app)
    app.mainloop()
