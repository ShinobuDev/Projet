import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import tkinter as ctk  # Pour une interface moderne
from PIL import Image, ImageTk
from bank_reconciliation import BankReconciliation

class BankReconciliationGUI:
    def __init__(self):
        # Configuration du thème
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Création de la fenêtre principale
        self.root = ctk.CTk()
        self.root.title("Rapprochement Bancaire - SOLIHA Normandie")
        self.root.geometry("800x600")
        
        # Chargement de l'icône
        try:
            icon_path = Path(__file__).parent / "assets" / "icon.ico"
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erreur lors du chargement de l'icône: {e}")
        
        # Couleurs SOLIHA
        self.colors = {
            'blue': '#005BBB',      # Bleu SOLIHA
            'light_blue': '#4D8DC9', # Bleu clair
            'white': '#FFFFFF',
            'gray': '#F5F5F5'
        }
        
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ctk.CTkFrame(self.root, fg_color=self.colors['white'])
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Logo SOLIHA avec gestion du chemin relatif
        try:
            logo_path = Path(__file__).parent / "assets" / "soliha_logo.png"
            logo_img = Image.open(logo_path)
            logo_img = logo_img.resize((200, 80), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(main_frame, image=self.logo, bg=self.colors['white'])
            logo_label.pack(pady=(0, 20))
        except Exception as e:
            print(f"Erreur lors du chargement du logo: {e}")
        
        # Titre
        title = ctk.CTkLabel(
            main_frame, 
            text="Rapprochement Bancaire",
            font=("Helvetica", 24, "bold"),
            text_color=self.colors['blue']
        )
        title.pack(pady=(0, 30))
        
        # Frame pour les fichiers
        files_frame = ctk.CTkFrame(main_frame, fg_color=self.colors['gray'])
        files_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        # Fichier 1
        file1_frame = self.create_file_input(
            files_frame, 
            "Fichier bancaire 1:",
            "Sélectionner le fichier 1"
        )
        file1_frame.pack(fill='x', padx=20, pady=10)
        
        # Fichier 2
        file2_frame = self.create_file_input(
            files_frame, 
            "Fichier bancaire 2:",
            "Sélectionner le fichier 2"
        )
        file2_frame.pack(fill='x', padx=20, pady=10)
        
        # Fichier de sortie
        output_frame = self.create_file_input(
            files_frame, 
            "Fichier de sortie:",
            "Sélectionner l'emplacement",
            is_save=True
        )
        output_frame.pack(fill='x', padx=20, pady=10)
        
        # Bouton de traitement
        self.process_button = ctk.CTkButton(
            main_frame,
            text="Lancer le rapprochement",
            font=("Helvetica", 14),
            command=self.process_files,
            fg_color=self.colors['blue'],
            hover_color=self.colors['light_blue']
        )
        self.process_button.pack(pady=30)
        
        # Barre de progression
        self.progress = ttk.Progressbar(
            main_frame,
            orient="horizontal",
            length=300,
            mode="determinate"
        )
        self.progress.pack(pady=(0, 20))
        
        # Zone de statut
        self.status_label = ctk.CTkLabel(
            main_frame,
            text="En attente...",
            font=("Helvetica", 12),
            text_color=self.colors['blue']
        )
        self.status_label.pack()

    def create_file_input(self, parent, label_text, button_text, is_save=False):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        
        label = ctk.CTkLabel(
            frame,
            text=label_text,
            font=("Helvetica", 12),
            text_color=self.colors['blue']
        )
        label.pack(side='left', padx=(0, 10))
        
        entry = ctk.CTkEntry(
            frame,
            width=300,
            font=("Helvetica", 12)
        )
        entry.pack(side='left', padx=(0, 10))
        
        button = ctk.CTkButton(
            frame,
            text=button_text,
            command=lambda: self.browse_file(entry, is_save),
            font=("Helvetica", 12),
            fg_color=self.colors['blue'],
            hover_color=self.colors['light_blue']
        )
        button.pack(side='left')
        
        return frame
        
    def browse_file(self, entry_widget, is_save=False):
        if is_save:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
        else:
            filename = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx")]
            )
            
        if filename:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filename)
            
    def process_files(self):
        # Simulation du traitement
        self.progress['value'] = 0
        self.status_label.configure(text="Traitement en cours...")
        
        # Récupération des chemins de fichiers
        file1 = self.file1_entry.get()
        file2 = self.file2_entry.get()
        output = self.output_entry.get()
        
        try:
            # Votre logique de traitement ici
            self.reconciliation = BankReconciliation()
            self.reconciliation.load_files(file1, file2)
            self.reconciliation.export_to_excel(output)
            
            self.progress['value'] = 100
            self.status_label.configure(text="Traitement terminé avec succès!")
            messagebox.showinfo("Succès", "Le rapprochement bancaire a été effectué avec succès!")
            
        except Exception as e:
            self.status_label.configure(text="Erreur lors du traitement!")
            messagebox.showerror("Erreur", str(e))
            
    def run(self):
        self.root.mainloop()