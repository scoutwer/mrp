"""ui_builder.py — Budowanie widgetów GUI (Tkinter)."""
import tkinter as tk
from tkinter import ttk


def setup_styles():
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TFrame", background="#f4f6f9")
    style.configure("Card.TFrame", background="#ffffff", relief="flat", borderwidth=1)
    style.configure("TLabel", background="#ffffff", font=("Segoe UI", 10), foreground="#333333")
    style.configure("Header.TLabel", background="#f4f6f9", font=("Segoe UI", 16, "bold"), foreground="#2c3e50")
    style.configure("Subheader.TLabel", background="#ffffff", font=("Segoe UI", 11, "bold"), foreground="#0056b3")
    style.configure("Export.TButton", font=("Segoe UI", 10, "bold"), background="#28a745", foreground="white", padding=8)
    style.map("Export.TButton", background=[("active", "#218838"), ("disabled", "#cccccc")])
    style.configure("SaveHistory.TButton", font=("Segoe UI", 10, "bold"), background="#007bff", foreground="white", padding=8)
    style.map("SaveHistory.TButton", background=[("active", "#0056b3")])
    style.configure("TNotebook", background="#f4f6f9", borderwidth=0)
    style.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=[15, 8], background="#e9ecef", foreground="#495057")
    style.map("TNotebook.Tab", background=[("selected", "#ffffff")], foreground=[("selected", "#0056b3")])


def build_magazyn_tab(tab, app):
    ttk.Label(tab, text="Baza Indeksów (Aktualne stany magazynowe)", style="Header.TLabel", background="#ffffff").pack(anchor="w", padx=20, pady=(20, 10))
    columns = ("nazwa", "typ", "zapas", "lt", "partia", "rodzic")
    app.tree_magazyn = ttk.Treeview(tab, columns=columns, show="headings", height=20)
    for cid, txt in [("nazwa","Nazwa Detalu / Indeks"),("typ","Typ Produktu"),("zapas","Zapas (Szt.)"),("lt","Czas Dostawy (Tyg.)"),("partia","Wielkość Partii"),("rodzic","Połączony Rodzic")]:
        app.tree_magazyn.heading(cid, text=txt)
    app.tree_magazyn.column("nazwa", width=250, anchor="w")
    app.tree_magazyn.column("typ", width=250, anchor="w")
    for c in ["zapas","lt","partia"]:
        app.tree_magazyn.column(c, width=120, anchor="center")
    app.tree_magazyn.column("rodzic", width=200, anchor="center")
    sb = ttk.Scrollbar(tab, orient="vertical", command=app.tree_magazyn.yview)
    app.tree_magazyn.configure(yscrollcommand=sb.set)
    app.tree_magazyn.pack(side="left", fill="both", expand=True, padx=(20, 0), pady=(0, 20))
    sb.pack(side="right", fill="y", padx=(0, 20), pady=(0, 20))


def build_rozrywka_tab(tab, app):
    import random
    center = tk.Frame(tab, bg="#ffffff")
    center.place(relx=0.5, rely=0.5, anchor="center")
    ttk.Label(center, text="\U0001f3b0 Kasyno Planisty", font=("Segoe UI", 24, "bold"), background="#ffffff", foreground="#2c3e50").pack(pady=(0, 10))
    ttk.Label(center, text="Zbyt wiele wyliczeń zapotrzebowania netto? Odpocznij!", font=("Segoe UI", 12), background="#ffffff").pack(pady=(0, 20))
    app.saldo_gry = 1000
    app.lbl_saldo = tk.Label(center, text=f"Twoje saldo: {app.saldo_gry} PLN", font=("Segoe UI", 16, "bold"), bg="#ffffff", fg="#28a745")
    app.lbl_saldo.pack(pady=10)
    app.slot_frame = tk.Frame(center, bg="#343a40", bd=10, relief="ridge")
    app.slot_frame.pack(pady=20)
    app.slots = []
    for i in range(3):
        lbl = tk.Label(app.slot_frame, text="❓", font=("Segoe UI", 60), bg="#f8f9fa", width=3, height=1, relief="sunken", bd=3)
        lbl.grid(row=0, column=i, padx=15, pady=15)
        app.slots.append(lbl)
    app.lbl_wynik = tk.Label(center, text="Zakręć bębnami, aby zagrać!", font=("Segoe UI", 14), bg="#ffffff", fg="#6c757d")
    app.lbl_wynik.pack(pady=10)
    app.btn_spin = tk.Button(center, text="ZAKRĘĆ (Koszt: 10 PLN)", bg="#ffc107", fg="#212529", font=("Segoe UI", 16, "bold"), cursor="hand2", relief="flat", padx=20, pady=10, command=app.spin_slots)
    app.btn_spin.pack(pady=20)
    app.btn_spin.bind("<Enter>", lambda e: e.widget.config(bg="#e0a800") if e.widget['state'] == tk.NORMAL else None)
    app.btn_spin.bind("<Leave>", lambda e: e.widget.config(bg="#ffc107") if e.widget['state'] == tk.NORMAL else None)
