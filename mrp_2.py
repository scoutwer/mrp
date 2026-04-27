import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import math
import json
import os
import random

class ReactiveMRPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("System MRP")
        self.root.geometry("1400x850")
        self.root.configure(bg="#f4f6f9")
        
        self.liczba_tygodni_var = tk.IntVar(value=10)
        
        self.plik_historii = "mrp_history.json"
        self.historia_planow = {}
        self.ostatnie_dane_mrp = None
        
        self.gross_entries = []
        self.receipt_entries = []
        
        self.setup_styles()
        self.build_ui()
        
        self.wczytaj_historie_z_pliku()
        self.process_mrp()

    def setup_styles(self):
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

    def build_ui(self):
        main_container = ttk.Frame(self.root, padding=15)
        main_container.pack(fill="both", expand=True)

        ttk.Label(main_container, text="Zintegrowany System MRP & Baza Magazynowa", style="Header.TLabel").pack(anchor="w", pady=(0, 10))

        # ==========================================
        # ZAKŁADKI (NOTEBOOK)
        # ==========================================
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill="both", expand=True)

        self.tab_mrp = ttk.Frame(self.notebook, style="TFrame")
        self.tab_magazyn = ttk.Frame(self.notebook, style="Card.TFrame")
        self.tab_rozrywka = ttk.Frame(self.notebook, style="Card.TFrame") # Nowa zakładka!

        self.notebook.add(self.tab_mrp, text="📊 Planowanie MRP")
        self.notebook.add(self.tab_magazyn, text="📦 Stan Magazynu")
        self.notebook.add(self.tab_rozrywka, text="🎰 Strefa Relaksu")

        self.build_mrp_tab()
        self.build_magazyn_tab()
        self.build_rozrywka_tab()

    def build_mrp_tab(self):
        content_frame = ttk.Frame(self.tab_mrp)
        content_frame.pack(fill="both", expand=True, pady=10)

        # ====== LEWY PANEL ======
        left_outer_frame = ttk.Frame(content_frame, style="Card.TFrame")
        left_outer_frame.pack(side="left", fill="y", padx=(0, 15))
        
        left_canvas = tk.Canvas(left_outer_frame, bg="#ffffff", highlightthickness=0, width=480)
        left_scrollbar = ttk.Scrollbar(left_outer_frame, orient="vertical", command=left_canvas.yview)
        left_panel = ttk.Frame(left_canvas, style="Card.TFrame", padding=15)
        
        left_panel.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.create_window((0, 0), window=left_panel, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")
        
        ttk.Label(left_panel, text="Wczytaj z bazy (Profil):", style="Subheader.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.combo_historia = ttk.Combobox(left_panel, state="readonly", width=25, font=("Segoe UI", 10))
        self.combo_historia.grid(row=0, column=1, pady=(0, 5), sticky="w")
        self.combo_historia.bind("<<ComboboxSelected>>", self.wczytaj_z_historii)
        
        params_header_frame = ttk.Frame(left_panel, style="Card.TFrame")
        params_header_frame.grid(row=1, column=0, columnspan=2, sticky="we", pady=(15, 10))
        ttk.Label(params_header_frame, text="Parametry Wejściowe", style="Subheader.TLabel").pack(side="left")
        
        self.btn_wyczysc = tk.Button(params_header_frame, text="⟲ Wyczyść", bg="#ffffff", fg="#6c757d", font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0, cursor="hand2", command=self.wyczysc_formularz)
        self.btn_wyczysc.pack(side="right")
        self.btn_wyczysc.bind("<Enter>", lambda e: e.widget.config(fg="#343a40", bg="#f8f9fa"))
        self.btn_wyczysc.bind("<Leave>", lambda e: e.widget.config(fg="#6c757d", bg="#ffffff"))

        labels = [
            ("Nazwa części / detalu:", ""),
            ("Na stanie (zapas początkowy):", ""),
            ("Czas realizacji (w tyg.):", ""),
            ("Wielkość partii (0=partia na partię):", "")
        ]
        
        self.entries = {}
        for i, (label_text, default_val) in enumerate(labels, start=2):
            ttk.Label(left_panel, text=label_text).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 10))
            entry = ttk.Entry(left_panel, width=24, font=("Segoe UI", 10))
            entry.insert(0, default_val)
            entry.grid(row=i, column=1, pady=5, sticky="w")
            entry.bind("<KeyRelease>", lambda event: self.process_mrp())
            self.entries[label_text] = entry

        ttk.Label(left_panel, text="Typ produktu (BOM):").grid(row=6, column=0, sticky="w", pady=5, padx=(0, 10))
        self.combo_typ_bom = ttk.Combobox(left_panel, state="readonly", width=22, font=("Segoe UI", 10))
        self.combo_typ_bom['values'] = ["Wyrób gotowy (Poziom 0)", "Półprodukt (Poziom 1)", "Podzespół (Poziom 2)", "Surowiec / Materiał (Poziom 3+)"]
        self.combo_typ_bom.set("Wyrób gotowy (Poziom 0)")
        self.combo_typ_bom.grid(row=6, column=1, pady=5, sticky="w")
        self.combo_typ_bom.bind("<<ComboboxSelected>>", lambda e: self.process_mrp())

        ttk.Label(left_panel, text="Zależności BOM (Opcjonalne)", style="Subheader.TLabel").grid(row=7, column=0, columnspan=2, sticky="w", pady=(20, 10))
        
        ttk.Label(left_panel, text="Produkt nadrzędny (Rodzic):").grid(row=8, column=0, sticky="w", pady=5)
        self.combo_rodzic = ttk.Combobox(left_panel, state="readonly", width=22, font=("Segoe UI", 10))
        self.combo_rodzic.grid(row=8, column=1, pady=5, sticky="w")
        self.combo_rodzic.bind("<<ComboboxSelected>>", lambda e: self.process_mrp())
        
        ttk.Label(left_panel, text="Ile sztuk na 1 rodzica:").grid(row=9, column=0, sticky="w", pady=5)
        self.entry_mnoznik = ttk.Entry(left_panel, width=24, font=("Segoe UI", 10))
        self.entry_mnoznik.insert(0, "1")
        self.entry_mnoznik.grid(row=9, column=1, pady=5, sticky="w")
        self.entry_mnoznik.bind("<KeyRelease>", lambda e: self.process_mrp())

        ttk.Label(left_panel, text="Horyzont planowania (tyg.):").grid(row=10, column=0, sticky="w", pady=(15, 5))
        spin_tygodnie = ttk.Spinbox(left_panel, from_=1, to=52, width=22, textvariable=self.liczba_tygodni_var, font=("Segoe UI", 10))
        spin_tygodnie.grid(row=10, column=1, pady=(15, 5), sticky="w")
        spin_tygodnie.bind("<<Increment>>", lambda e: self.rebuild_grid())
        spin_tygodnie.bind("<<Decrement>>", lambda e: self.rebuild_grid())
        spin_tygodnie.bind("<KeyRelease>", lambda e: self.rebuild_grid())

        ttk.Label(left_panel, text="Planowanie okresowe:", style="Subheader.TLabel").grid(row=11, column=0, columnspan=2, sticky="w", pady=(20, 10))
        
        self.input_scroll_container = ttk.Frame(left_panel, style="Card.TFrame")
        self.input_scroll_container.grid(row=12, column=0, columnspan=2, sticky="we")

        self.input_canvas = tk.Canvas(self.input_scroll_container, bg="#ffffff", highlightthickness=0, height=105, width=420)
        self.input_scrollbar = ttk.Scrollbar(self.input_scroll_container, orient="horizontal", command=self.input_canvas.xview)

        self.grid_wrapper = ttk.Frame(self.input_canvas, style="Card.TFrame")
        self.grid_wrapper.bind("<Configure>", lambda e: self.input_canvas.configure(scrollregion=self.input_canvas.bbox("all")))

        self.input_canvas.create_window((0, 0), window=self.grid_wrapper, anchor="nw")
        self.input_canvas.configure(xscrollcommand=self.input_scrollbar.set)

        self.input_canvas.pack(side="top", fill="x", expand=True)
        self.input_scrollbar.pack(side="bottom", fill="x")

        self.rebuild_grid(initial=True)

        btn_frame = ttk.Frame(left_panel, style="Card.TFrame")
        btn_frame.grid(row=13, column=0, columnspan=2, sticky="we", pady=(25, 0))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        
        self.btn_zapisz = tk.Button(btn_frame, text="💾 ZAPISZ PROFIL", bg="#cccccc", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", state=tk.DISABLED, cursor="arrow", command=self.zapisz_do_historii)
        self.btn_zapisz.grid(row=0, column=0, sticky="we", padx=(0, 5), ipady=8)
        self.btn_zapisz.bind("<Enter>", lambda e: e.widget.config(bg="#0056b3") if e.widget['state'] == tk.NORMAL else None)
        self.btn_zapisz.bind("<Leave>", lambda e: e.widget.config(bg="#007bff") if e.widget['state'] == tk.NORMAL else None)
        
        self.btn_export = tk.Button(btn_frame, text="📊 EXCEL", bg="#cccccc", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", state=tk.DISABLED, cursor="arrow", command=self.export_excel)
        self.btn_export.grid(row=0, column=1, sticky="we", padx=(5, 0), ipady=8)
        self.btn_export.bind("<Enter>", lambda e: e.widget.config(bg="#218838") if e.widget['state'] == tk.NORMAL else None)
        self.btn_export.bind("<Leave>", lambda e: e.widget.config(bg="#28a745") if e.widget['state'] == tk.NORMAL else None)

        # ====== PRAWY PANEL ======
        right_panel = ttk.Frame(content_frame, style="Card.TFrame", padding=20)
        right_panel.pack(side="right", fill="both", expand=True)
        
        self.alert_frame = ttk.Frame(right_panel, style="Card.TFrame")
        self.alert_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(self.alert_frame, text="Status systemu:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.alert_text = tk.Text(self.alert_frame, height=3, bg="#f8f9fa", fg="#333333", font=("Segoe UI", 10, "bold"), relief="flat", padx=10, pady=10)
        self.alert_text.pack(fill="x", pady=(5, 0))
        self.alert_text.config(state="disabled")

        table_header_frame = ttk.Frame(right_panel, style="Card.TFrame")
        table_header_frame.pack(fill="x", pady=(10, 10))
        
        ttk.Label(table_header_frame, text="Wygenerowana Tabela MRP", style="Subheader.TLabel").pack(side="left")
        
        self.btn_usun = tk.Button(table_header_frame, text="🗑 Usuń z bazy", bg="#ffffff", fg="#dc3545", font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0, cursor="hand2", command=self.usun_z_historii)
        self.btn_usun.pack(side="right")
        self.btn_usun.bind("<Enter>", lambda e: e.widget.config(fg="#a71d2a", bg="#f8d7da"))
        self.btn_usun.bind("<Leave>", lambda e: e.widget.config(fg="#dc3545", bg="#ffffff"))

        self.table_container = ttk.Frame(right_panel)
        self.table_container.pack(fill="both", expand=True)

        self.canvas_mrp = tk.Canvas(self.table_container, bg="#ffffff", highlightthickness=0)
        self.scrollbar_mrp = ttk.Scrollbar(self.table_container, orient="horizontal", command=self.canvas_mrp.xview)
        
        self.scrollable_frame = tk.Frame(self.canvas_mrp, bg="#adb5bd") 
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas_mrp.configure(scrollregion=self.canvas_mrp.bbox("all")))

        self.canvas_mrp.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas_mrp.configure(xscrollcommand=self.scrollbar_mrp.set)

        self.canvas_mrp.pack(side="top", fill="both", expand=True)
        self.scrollbar_mrp.pack(side="bottom", fill="x")

    def build_magazyn_tab(self):
        ttk.Label(self.tab_magazyn, text="Baza Indeksów (Aktualne stany magazynowe)", style="Header.TLabel", background="#ffffff").pack(anchor="w", padx=20, pady=(20, 10))
        
        columns = ("nazwa", "typ", "zapas", "lt", "partia", "rodzic")
        self.tree_magazyn = ttk.Treeview(self.tab_magazyn, columns=columns, show="headings", height=20)
        
        self.tree_magazyn.heading("nazwa", text="Nazwa Detalu / Indeks")
        self.tree_magazyn.heading("typ", text="Typ Produktu")
        self.tree_magazyn.heading("zapas", text="Zapas (Szt.)")
        self.tree_magazyn.heading("lt", text="Czas Dostawy (Tyg.)")
        self.tree_magazyn.heading("partia", text="Wielkość Partii")
        self.tree_magazyn.heading("rodzic", text="Połączony Rodzic")
        
        self.tree_magazyn.column("nazwa", width=250, anchor="w")
        self.tree_magazyn.column("typ", width=250, anchor="w")
        self.tree_magazyn.column("zapas", width=120, anchor="center")
        self.tree_magazyn.column("lt", width=150, anchor="center")
        self.tree_magazyn.column("partia", width=120, anchor="center")
        self.tree_magazyn.column("rodzic", width=200, anchor="center")
        
        scrollbar = ttk.Scrollbar(self.tab_magazyn, orient="vertical", command=self.tree_magazyn.yview)
        self.tree_magazyn.configure(yscrollcommand=scrollbar.set)
        
        self.tree_magazyn.pack(side="left", fill="both", expand=True, padx=(20, 0), pady=(0, 20))
        scrollbar.pack(side="right", fill="y", padx=(0, 20), pady=(0, 20))

    # ==========================================
    # EASTER EGG: STREFA RELAKSU (JEDNORĘKI BANDYTA)
    # ==========================================
    def build_rozrywka_tab(self):
        # Główny kontener wyśrodkowany
        center_frame = tk.Frame(self.tab_rozrywka, bg="#ffffff")
        center_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        ttk.Label(center_frame, text="🎰 Kasyno Planisty", font=("Segoe UI", 24, "bold"), background="#ffffff", foreground="#2c3e50").pack(pady=(0, 10))
        ttk.Label(center_frame, text="Zbyt wiele wyliczeń zapotrzebowania netto? Odpocznij!", font=("Segoe UI", 12), background="#ffffff").pack(pady=(0, 20))

        self.saldo_gry = 1000
        self.lbl_saldo = tk.Label(center_frame, text=f"Twoje saldo: {self.saldo_gry} PLN", font=("Segoe UI", 16, "bold"), bg="#ffffff", fg="#28a745")
        self.lbl_saldo.pack(pady=10)

        # Maszyna
        self.slot_frame = tk.Frame(center_frame, bg="#343a40", bd=10, relief="ridge")
        self.slot_frame.pack(pady=20)

        self.slots = []
        for i in range(3):
            lbl = tk.Label(self.slot_frame, text="❓", font=("Segoe UI", 60), bg="#f8f9fa", width=3, height=1, relief="sunken", bd=3)
            lbl.grid(row=0, column=i, padx=15, pady=15)
            self.slots.append(lbl)

        self.lbl_wynik = tk.Label(center_frame, text="Zakręć bębnami, aby zagrać!", font=("Segoe UI", 14), bg="#ffffff", fg="#6c757d")
        self.lbl_wynik.pack(pady=10)

        self.btn_spin = tk.Button(center_frame, text="ZAKRĘĆ (Koszt: 10 PLN)", bg="#ffc107", fg="#212529", font=("Segoe UI", 16, "bold"), cursor="hand2", relief="flat", padx=20, pady=10, command=self.spin_slots)
        self.btn_spin.pack(pady=20)
        self.btn_spin.bind("<Enter>", lambda e: e.widget.config(bg="#e0a800") if e.widget['state'] == tk.NORMAL else None)
        self.btn_spin.bind("<Leave>", lambda e: e.widget.config(bg="#ffc107") if e.widget['state'] == tk.NORMAL else None)

    def spin_slots(self):
        if self.saldo_gry < 10:
            self.lbl_wynik.config(text="Brak środków! Czas wrócić do planowania MRP! 😅", fg="#dc3545")
            return

        self.saldo_gry -= 10
        self.lbl_saldo.config(text=f"Twoje saldo: {self.saldo_gry} PLN")
        self.btn_spin.config(state=tk.DISABLED, bg="#cccccc", cursor="arrow")
        self.lbl_wynik.config(text="Maszyna losuje...", fg="#007bff")

        self.spin_count = 0
        self.animate_spin()

    def animate_spin(self):
        symbole = ["🍒", "🍋", "🍉", "🔔", "💎", "🍀", "7️⃣"]
        if self.spin_count < 15:
            for lbl in self.slots:
                lbl.config(text=random.choice(symbole))
            self.spin_count += 1
            self.root.after(60, self.animate_spin) # Animacja co 60ms
        else:
            # Finał
            wyniki = [random.choice(symbole) for _ in range(3)]
            for i, lbl in enumerate(self.slots):
                lbl.config(text=wyniki[i])

            # Sprawdzanie wygranej
            if wyniki[0] == wyniki[1] == wyniki[2]:
                wygrana = 500 if wyniki[0] == "7️⃣" else 100
                self.saldo_gry += wygrana
                self.lbl_saldo.config(text=f"Twoje saldo: {self.saldo_gry} PLN")
                self.lbl_wynik.config(text=f"🎉 JACKPOT! Wygrywasz {wygrana} PLN! 🎉", fg="#28a745")
            elif wyniki[0] == wyniki[1] or wyniki[1] == wyniki[2] or wyniki[0] == wyniki[2]:
                # Dwa takie same - mała nagroda pocieszenia
                self.saldo_gry += 5
                self.lbl_saldo.config(text=f"Twoje saldo: {self.saldo_gry} PLN")
                self.lbl_wynik.config(text="Dwa symbole! Zwracamy 5 PLN.", fg="#17a2b8")
            else:
                self.lbl_wynik.config(text="Brak wygranej. Spróbuj ponownie!", fg="#6c757d")

            self.btn_spin.config(state=tk.NORMAL, bg="#ffc107", cursor="hand2")

    # ==========================================
    # LOGIKA SYSTEMU MRP
    # ==========================================
    def wczytaj_historie_z_pliku(self):
        if os.path.exists(self.plik_historii):
            try:
                with open(self.plik_historii, 'r', encoding='utf-8') as f:
                    self.historia_planow = json.load(f)
                self.aktualizuj_listy_wyboru()
                self.odswiez_tabele_magazynu()
            except Exception:
                pass

    def odswiez_tabele_magazynu(self):
        for row in self.tree_magazyn.get_children():
            self.tree_magazyn.delete(row)
            
        for k, v in self.historia_planow.items():
            nazwa = v.get("nazwa", "Brak")
            typ = v.get("typ_bom", "Nieznany")
            zapas = v.get("zapas", 0)
            lt = v.get("lt", 0)
            partia = v.get("partia", 0)
            if partia == 0: partia = "Lot-for-Lot"
            rodzic = v.get("rodzic", "-")
            if not rodzic: rodzic = "-"
            
            self.tree_magazyn.insert("", "end", values=(nazwa, typ, zapas, lt, partia, rodzic))

    def aktualizuj_listy_wyboru(self):
        klucze = list(self.historia_planow.keys())
        self.combo_historia['values'] = klucze
        self.combo_rodzic['values'] = [""] + klucze

    def zapisz_historie_do_pliku(self):
        try:
            with open(self.plik_historii, 'w', encoding='utf-8') as f:
                json.dump(self.historia_planow, f, indent=4, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("Błąd zapisu", f"Nie udało się trwale zapisać historii:\n{e}")

    def wyczysc_formularz(self):
        self.combo_historia.set('')
        self.combo_rodzic.set('')
        self.combo_typ_bom.set("Wyrób gotowy (Poziom 0)")
        self.entry_mnoznik.delete(0, tk.END)
        self.entry_mnoznik.insert(0, "1")
        
        for entry in self.entries.values():
            entry.delete(0, tk.END)
            
        self.liczba_tygodni_var.set(10)
        self.rebuild_grid(initial=True) 
        self.process_mrp() 

    def usun_z_historii(self):
        nazwa = self.combo_historia.get()
        if not nazwa or nazwa not in self.historia_planow:
            messagebox.showwarning("Brak wyboru", "Wybierz profil do usunięcia.")
            return

        if messagebox.askyesno("Potwierdzenie", f"Czy usunąć '{nazwa}' z bazy?"):
            if messagebox.askyesno("Ostateczne usunięcie", "Tej operacji NIE BĘDZIE można cofnąć.\nJesteś pewien?"):
                del self.historia_planow[nazwa]
                self.zapisz_historie_do_pliku()
                self.aktualizuj_listy_wyboru()
                self.odswiez_tabele_magazynu()
                self.wyczysc_formularz()

    def rebuild_grid(self, initial=False, dane_z_historii=None, *args):
        try:
            n = self.liczba_tygodni_var.get()
            if n < 1 or n > 52: return
        except tk.TclError:
            return

        if dane_z_historii:
            stare_gross = [str(x) if x != 0 else "" for x in dane_z_historii['gross']]
            stare_receipts = [str(x) if x != 0 else "" for x in dane_z_historii['sched']]
        else:
            stare_gross = [e.get().strip() for e in self.gross_entries] if not initial else []
            stare_receipts = [e.get().strip() for e in self.receipt_entries] if not initial else []

        for widget in self.grid_wrapper.winfo_children():
            widget.destroy()

        self.gross_entries.clear()
        self.receipt_entries.clear()

        ttk.Label(self.grid_wrapper, text="Tydzień:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="e", padx=5, pady=5)
        ttk.Label(self.grid_wrapper, text="Całkowite zapotrzeb.:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.grid_wrapper, text="Planowane przyjęcia:").grid(row=2, column=0, sticky="e", padx=5, pady=2)

        for i in range(n):
            ttk.Label(self.grid_wrapper, text=str(i+1), font=("Segoe UI", 9, "bold")).grid(row=0, column=i+1, pady=5)
            
            val_g = stare_gross[i] if i < len(stare_gross) else ""
            val_r = stare_receipts[i] if i < len(stare_receipts) else ""

            g_entry = ttk.Entry(self.grid_wrapper, width=4, justify="center")
            g_entry.insert(0, val_g)
            g_entry.grid(row=1, column=i+1, padx=2, pady=2)
            g_entry.bind("<KeyRelease>", lambda event: self.process_mrp())
            self.gross_entries.append(g_entry)
            
            r_entry = ttk.Entry(self.grid_wrapper, width=4, justify="center")
            r_entry.insert(0, val_r)
            r_entry.grid(row=2, column=i+1, padx=2, pady=2)
            r_entry.bind("<KeyRelease>", lambda event: self.process_mrp())
            self.receipt_entries.append(r_entry)
            
        if not initial:
            self.process_mrp()

    def pobierz_liste_z_pól(self, lista_entry):
        wynik = []
        for entry in lista_entry:
            val = entry.get().strip()
            try:
                wynik.append(int(val) if val else 0)
            except ValueError:
                wynik.append(0) 
        return wynik

    def zapisz_do_historii(self):
        if not self.ostatnie_dane_mrp:
            return
            
        nazwa = self.entries["Nazwa części / detalu:"].get().strip()
        if not nazwa:
            messagebox.showwarning("Brak nazwy", "Wpisz nazwę detalu przed zapisaniem.")
            return

        self.ostatnie_dane_mrp["nazwa"] = nazwa
        self.historia_planow[nazwa] = self.ostatnie_dane_mrp
        self.zapisz_historie_do_pliku()
        
        self.aktualizuj_listy_wyboru()
        self.combo_historia.set(nazwa)
        self.odswiez_tabele_magazynu() 
        
        messagebox.showinfo("Zapisano", f"Plan dla: '{nazwa}' został zachowany w Bazie Magazynowej.")

    def wczytaj_z_historii(self, event=None):
        wybrana_nazwa = self.combo_historia.get()
        if wybrana_nazwa not in self.historia_planow:
            return
            
        dane = self.historia_planow[wybrana_nazwa]
        
        self.entries["Nazwa części / detalu:"].delete(0, tk.END)
        self.entries["Nazwa części / detalu:"].insert(0, dane['nazwa'])
        
        self.entries["Na stanie (zapas początkowy):"].delete(0, tk.END)
        self.entries["Na stanie (zapas początkowy):"].insert(0, str(dane['zapas']) if dane['zapas'] != 0 else "")
        
        self.entries["Czas realizacji (w tyg.):"].delete(0, tk.END)
        self.entries["Czas realizacji (w tyg.):"].insert(0, str(dane['lt']) if dane['lt'] != 0 else "")
        
        self.entries["Wielkość partii (0=partia na partię):"].delete(0, tk.END)
        self.entries["Wielkość partii (0=partia na partię):"].insert(0, str(dane['partia']) if dane['partia'] != 0 else "")
        
        self.combo_typ_bom.set(dane.get('typ_bom', "Wyrób gotowy (Poziom 0)"))
        self.combo_rodzic.set(dane.get('rodzic', ''))
        self.entry_mnoznik.delete(0, tk.END)
        self.entry_mnoznik.insert(0, str(dane.get('mnoznik', 1)))
        
        self.liczba_tygodni_var.set(dane['n'])
        self.rebuild_grid(dane_z_historii=dane)

    def process_mrp(self):
        try:
            nazwa = self.entries["Nazwa części / detalu:"].get().strip()
            
            def bezpieczny_int(pole_nazwa):
                val = self.entries[pole_nazwa].get().strip()
                return int(val) if val else 0
                
            na_stanie = bezpieczny_int("Na stanie (zapas początkowy):")
            czas_realizacji = bezpieczny_int("Czas realizacji (w tyg.):")
            wielkosc_partii = bezpieczny_int("Wielkość partii (0=partia na partię):")
            
            typ_bom = self.combo_typ_bom.get()
            rodzic = self.combo_rodzic.get().strip()
            mnoznik_str = self.entry_mnoznik.get().strip()
            mnoznik = int(mnoznik_str) if mnoznik_str.isdigit() else 1
            n = self.liczba_tygodni_var.get()

            czy_dziala_bom = False
            if rodzic and rodzic in self.historia_planow and rodzic != nazwa:
                czy_dziala_bom = True
                dane_rodzica = self.historia_planow[rodzic]
                zamowienia_rodzica = dane_rodzica.get('rel', [])
                
                for i, entry in enumerate(self.gross_entries):
                    entry.config(state=tk.NORMAL)
                    entry.delete(0, tk.END)
                    
                    zapotrzebowanie = 0
                    if i < len(zamowienia_rodzica):
                        zapotrzebowanie = zamowienia_rodzica[i] * mnoznik
                        
                    entry.insert(0, str(zapotrzebowanie) if zapotrzebowanie > 0 else "")
                    entry.config(state="readonly")
            else:
                for entry in self.gross_entries:
                    if entry.cget("state") == "readonly":
                        entry.config(state=tk.NORMAL)

            gross_req = self.pobierz_liste_z_pól(self.gross_entries)
            sched_rec = self.pobierz_liste_z_pól(self.receipt_entries)
            
            if n == 0: return

        except Exception:
            return 

        proj_avail = [0] * n
        net_req = [0] * n
        planned_order_rec = [0] * n
        planned_order_rel = [0] * n
        bledy = []

        czy_pusty_plan = (sum(gross_req) == 0 and na_stanie == 0 and sum(sched_rec) == 0)

        if czy_pusty_plan or not nazwa:
            self.btn_zapisz.config(state=tk.DISABLED, bg="#cccccc", cursor="arrow")
            self.btn_export.config(state=tk.DISABLED, bg="#cccccc", cursor="arrow")
        else:
            self.btn_zapisz.config(state=tk.NORMAL, bg="#007bff", cursor="hand2")
            self.btn_export.config(state=tk.NORMAL, bg="#28a745", cursor="hand2")

        for i in range(n):
            zapas_poprzedni = proj_avail[i-1] if i > 0 else na_stanie
            stan_przed = zapas_poprzedni + sched_rec[i] - gross_req[i]
            
            if stan_przed < 0:
                brakuje = abs(stan_przed)
                net_req[i] = brakuje
                
                if wielkosc_partii > 0:
                    mnozenie = math.ceil(brakuje / wielkosc_partii)
                    zamowienie = mnozenie * wielkosc_partii
                else:
                    zamowienie = brakuje
                    
                rel_idx = i - czas_realizacji
                
                if rel_idx >= 0:
                    planned_order_rel[rel_idx] = zamowienie
                    planned_order_rec[i] = zamowienie
                    proj_avail[i] = stan_przed + zamowienie
                else:
                    bledy.append(f"⚠ BŁĄD LOGISTYCZNY (Tydzień {i+1}): Brak czasu na zamówienie! Zapas spada do {stan_przed} szt.")
                    planned_order_rec[i] = 0
                    proj_avail[i] = stan_przed 
            else:
                net_req[i] = 0
                planned_order_rec[i] = 0
                proj_avail[i] = stan_przed

        self.alert_text.config(state="normal")
        self.alert_text.delete("1.0", tk.END)
        
        if czy_pusty_plan:
             self.alert_text.config(bg="#f8f9fa", fg="#333333") 
             self.alert_text.insert(tk.END, "Oczekuję na wprowadzenie danych do planu...")
        elif bledy:
            self.alert_text.config(bg="#f8d7da", fg="#721c24") 
            self.alert_text.insert(tk.END, "\n".join(bledy))
        else:
            self.alert_text.config(bg="#d4edda", fg="#155724") 
            if czy_dziala_bom:
                self.alert_text.insert(tk.END, f"✓ Plan optymalny. Zapotrzebowanie pobrane dynamicznie z: '{rodzic}'.")
            else:
                self.alert_text.insert(tk.END, "✓ Plan optymalny. Zapas pokrywa zapotrzebowanie, dostawy o czasie.")
                
        self.alert_text.config(state="disabled")

        self.update_table_grid(gross_req, sched_rec, proj_avail, net_req, planned_order_rel, planned_order_rec, n)
        
        self.ostatnie_dane_mrp = {
            "nazwa": nazwa, "lt": czas_realizacji, "partia": wielkosc_partii, "typ_bom": typ_bom, "zapas": na_stanie,
            "rodzic": rodzic, "mnoznik": mnoznik,
            "gross": gross_req, "sched": sched_rec, "avail": proj_avail, "net": net_req, "rel": planned_order_rel, "rec": planned_order_rec, "n": n
        }

    def update_table_grid(self, gross, sched, avail, net, rel, rec, n):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        tk.Label(self.scrollable_frame, text="Dane produkcyjne", font=("Segoe UI", 10, "bold"), bg="#dce4ec", width=25, anchor="w", padx=10, pady=8).grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        for i in range(n):
            tk.Label(self.scrollable_frame, text=str(i+1), font=("Segoe UI", 10, "bold"), bg="#dce4ec", width=6, pady=8).grid(row=0, column=i+1, sticky="nsew", padx=1, pady=1)

        wiersze = [
            ("Całkowite zapotrzebowanie", gross),
            ("Planowane przyjęcia", sched),
            ("Przewidywane na stanie", avail),
            ("Zapotrzebowanie netto", net),
            ("Planowane zamówienia", rel),
            ("Planowane przyjęcie zamówień", rec)
        ]

        for r, (nazwa, wartosci) in enumerate(wiersze, start=1):
            for c, w in enumerate(wartosci, start=1):
                bg_nazwy = "#f8f9fa"
                bg_wartosci = "#ffffff"
                fg_color = "#333333"
                fnt = ("Segoe UI", 10)
                
                if nazwa == "Przewidywane na stanie":
                    bg_nazwy = "#e2eef9"
                    fnt = ("Segoe UI", 10, "bold")
                    if w < 0:
                        bg_wartosci = "#f8d7da"
                        fg_color = "#dc3545"
                    else:
                        bg_wartosci = "#e2eef9"
                        fg_color = "#0056b3"
                    val = str(w) 
                else:
                    val = str(w) if w > 0 else "" 

                if c == 1:
                    tk.Label(self.scrollable_frame, text=nazwa, font=fnt, bg=bg_nazwy, fg="#333333" if nazwa!="Przewidywane na stanie" else "#0056b3", anchor="w", padx=10, pady=8).grid(row=r, column=0, sticky="nsew", padx=1, pady=1)
                
                tk.Label(self.scrollable_frame, text=val, font=fnt, bg=bg_wartosci, fg=fg_color, anchor="center", pady=8).grid(row=r, column=c, sticky="nsew", padx=1, pady=1)

    def export_excel(self):
        d = self.ostatnie_dane_mrp
        if not d: return

        dane_excel = {
            "Dane produkcyjne": [
                "Całkowite zapotrzebowanie",
                "Planowane przyjęcia",
                "Przewidywane na stanie",
                "Zapotrzebowanie netto",
                "Planowane zamówienia",
                "Planowane przyjęcie zamówień"
            ]
        }
        
        for i in range(d["n"]):
            dane_excel[i+1] = [
                d["gross"][i] if d["gross"][i] > 0 else "",
                d["sched"][i] if d["sched"][i] > 0 else "",
                d["avail"][i],
                d["net"][i] if d["net"][i] > 0 else "",
                d["rel"][i] if d["rel"][i] > 0 else "",
                d["rec"][i] if d["rec"][i] > 0 else ""
            ]

        df = pd.DataFrame(dane_excel)
        nazwa_pliku_bezpieczna = d['nazwa'].replace(' ', '_')
        if nazwa_pliku_bezpieczna == "":
            nazwa_pliku_bezpieczna = "Bez_Nazwy"
            
        filename = f"Analiza_MRP_{nazwa_pliku_bezpieczna}.xlsx"

        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                sheet_name = d["nazwa"][:31] if d["nazwa"] else "Arkusz1"
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                ws = writer.sheets[sheet_name]
                
                stopka_wiersz = len(df) + 2
                ws.cell(row=stopka_wiersz, column=1, value=f"Czas realizacji = {d['lt']}")
                ws.cell(row=stopka_wiersz, column=2, value=f"Wielkość partii = {d['partia']}")
                ws.cell(row=stopka_wiersz, column=3, value=f"Typ: {d.get('typ_bom')}")
                ws.cell(row=stopka_wiersz, column=4, value=f"Na stanie = {d['zapas']}")
                
                ws.cell(row=stopka_wiersz+1, column=1, value=f"Rodzic BOM: {d.get('rodzic') or 'Brak'}")
                ws.cell(row=stopka_wiersz+1, column=2, value=f"Sztuk na 1 rodzica: {d.get('mnoznik', 1)}")
                
                ws.column_dimensions['A'].width = 30
                for col in range(2, d["n"]+2):
                    ws.column_dimensions[chr(64+col)].width = 8

            messagebox.showinfo("Eksport Zakończony", f"Plik zapisany poprawnie jako:\n{filename}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Zamknij plik Excel, jeśli jest aktualnie otwarty.\nSzczegóły: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReactiveMRPApp(root)
    root.mainloop()