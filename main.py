"""main.py — Punkt wejścia aplikacji MRP."""
import tkinter as tk
from gui import ReactiveMRPApp

if __name__ == "__main__":
    root = tk.Tk()
    app = ReactiveMRPApp(root)
    root.mainloop()
