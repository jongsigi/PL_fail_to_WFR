# main.py
import tkinter as tk
from PathSelector.FolderSelectorApp import FolderSelectorApp

# GUI 실행
if __name__ == "__main__":
    root = tk.Tk()
    app = FolderSelectorApp(root)
    root.mainloop()