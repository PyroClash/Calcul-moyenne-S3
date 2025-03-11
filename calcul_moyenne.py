import sys
import subprocess
import importlib.util

def check_and_install_packages():
    required_packages = ['tkinter', 'PyPDF2']
    
    for package in required_packages:
        # tkinter est généralement inclus avec Python, pas besoin de l'installer via pip
        if package == 'tkinter':
            continue
            
        # Vérifie si le package est installé
        if importlib.util.find_spec(package) is None:
            print(f"Installation du module {package}...")
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
                print(f"{package} a été installé avec succès!")
            except subprocess.CalledProcessError as e:
                print(f"Erreur lors de l'installation de {package}: {e}")
                sys.exit(1)

# Vérifier et installer les packages au démarrage
check_and_install_packages()

import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PyPDF2 import PdfReader
# Importer les matières et les coefficients depuis le nouveau fichier
from matieres_coeffs import MATIERES, COEFFICIENTS

class ApplicationNotes:
    def __init__(self, root):
        self.root = root
        self.root.title("Calcul des moyennes UE - Made by PDO")
        
        # Dictionnaire pour stocker les notes de chaque matière
        self.notes_par_matiere = {}
        
        # Frame principal
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame pour le chargement PDF
        self.pdf_frame = ttk.LabelFrame(self.main_frame, text="Chargement des notes", padding="5")
        self.pdf_frame.pack(fill=tk.X, pady=5)
        
        # Bouton pour charger un relevé PDF
        ttk.Button(self.pdf_frame, text="Charger un relevé PDF", 
                  command=self.charger_pdf).pack(side=tk.LEFT, padx=5, pady=10)
        
        # Checkbox pour ignorer SAE3.01
        self.ignore_sae = tk.BooleanVar(value=True)  # Activée par défaut
        ttk.Checkbutton(self.pdf_frame, text="Ignorer et neutraliser SAE3.01", 
                        variable=self.ignore_sae).pack(side=tk.RIGHT, padx=5, pady=10)
        
        # Frame pour l'affichage des notes
        self.notes_frame = ttk.LabelFrame(self.main_frame, text="Notes chargées", padding="5")
        self.notes_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Liste des notes
        self.notes_text = tk.Text(self.notes_frame, height=10, width=50)
        self.notes_text.pack(fill=tk.BOTH, expand=True)
        
        # Boutons de contrôle
        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(self.control_frame, text="Calculer les moyennes", 
                  command=self.calculer_moyennes).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.control_frame, text="Effacer tout", 
                  command=self.effacer_tout).pack(side=tk.RIGHT, padx=5)
        
        # Zone de résultats
        self.resultats_frame = ttk.LabelFrame(self.main_frame, text="Résultats", padding="5")
        self.resultats_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.resultats = tk.Text(self.resultats_frame, height=8, width=50)
        self.resultats.pack(fill=tk.BOTH, expand=True)
    
    def afficher_notes(self):
        self.notes_text.delete(1.0, tk.END)
        for matiere in sorted(self.notes_par_matiere.keys()):
            self.notes_text.insert(tk.END, f"\n{matiere}:\n")
            for i, (note, coeff) in enumerate(self.notes_par_matiere[matiere], 1):
                self.notes_text.insert(tk.END, 
                    f"  Note {i}: {note:.2f}/20 (coeff: {coeff:.2f})\n")
    
    def effacer_tout(self):
        if messagebox.askyesno("Confirmation", "Voulez-vous vraiment effacer toutes les notes ?"):
            self.notes_par_matiere.clear()
            self.notes_text.delete(1.0, tk.END)
            self.resultats.delete(1.0, tk.END)
    
    def calculer_moyennes(self):
        # Utiliser les coefficients importés
        coeffs = COEFFICIENTS
        
        # Convertir les clés des notes au bon format (R3.301 -> R3.01)
        notes_converties = {}
        for matiere, notes in self.notes_par_matiere.items():
            if matiere == 'SAE3.01':  # Pas besoin de convertir les SAE
                if not self.ignore_sae.get():  # Ajouter seulement si on n'ignore pas SAE3.01
                    notes_converties[matiere] = notes
            else:
                # Extraire le numéro à 3 chiffres et le convertir en 2 chiffres
                num = matiere.split('.')[1]  # '301', '302', etc.
                nouvelle_cle = f"R3.{int(num[-2:]):02d}"  # Prend les 2 derniers chiffres
                notes_converties[nouvelle_cle] = notes
        
        print("\nNotes après conversion des clés:")
        for matiere, notes in sorted(notes_converties.items()):
            print(f"{matiere}: {notes}")
        
        # Calcul et affichage des moyennes par UE
        self.resultats.delete(1.0, tk.END)
        self.resultats.insert(tk.END, "Moyennes par UE :\n")
        self.resultats.insert(tk.END, "-" * 40 + "\n")
        
        for ue, matieres_coeffs in coeffs.items():
            print(f"\nCalcul pour {ue}:")
            somme_ponderee = 0
            somme_coeffs = 0
            
            for matiere, coeff_ue in matieres_coeffs.items():
                print(f"  Traitement de {matiere} (coeff UE: {coeff_ue})")
                
                # Ignorer SAE3.01 si la case est cochée
                if matiere == 'SAE3.01' and self.ignore_sae.get():
                    print(f"    SAE3.01 ignorée car la case est cochée")
                    continue
                    
                if matiere in notes_converties:
                    notes_matiere = notes_converties[matiere]
                    print(f"    Notes trouvées: {notes_matiere}")
                    
                    # Calculer la moyenne pondérée de la matière
                    somme_ponderee_matiere = sum(note * coeff for note, coeff in notes_matiere)
                    somme_coeffs_matiere = sum(coeff for _, coeff in notes_matiere)
                    
                    if somme_coeffs_matiere > 0:
                        moyenne_matiere = somme_ponderee_matiere / somme_coeffs_matiere
                        contribution = moyenne_matiere * coeff_ue
                        somme_ponderee += contribution
                        somme_coeffs += coeff_ue
                        print(f"    Moyenne matière: {moyenne_matiere:.2f}")
                        print(f"    Contribution à l'UE: {contribution:.2f}")
                else:
                    print(f"    Pas de notes pour cette matière")
            
            print(f"  Somme pondérée finale: {somme_ponderee:.2f}")
            print(f"  Somme coeffs finale: {somme_coeffs:.2f}")
            
            if somme_coeffs > 0:
                moyenne = somme_ponderee / somme_coeffs
                self.resultats.insert(tk.END, f"{ue}: {moyenne:.2f}/20\n")
                print(f"  Moyenne {ue}: {moyenne:.2f}/20")
            else:
                self.resultats.insert(tk.END, f"{ue}: Notes manquantes\n")
                print(f"  {ue}: Notes manquantes")

    def charger_pdf(self):
        # Ouvrir le sélecteur de fichier
        pdf_path = filedialog.askopenfilename(
            title="Sélectionner le relevé de notes PDF",
            filetypes=[("Fichiers PDF", "*.pdf")]
        )
        
        if pdf_path:
            # Convertir le PDF en texte
            texte = pdf_to_text(pdf_path)
            if texte:
                # Extraire les notes
                notes = extraire_notes_from_txt(texte)
                
                # Mettre à jour les notes dans l'application
                self.notes_par_matiere = notes
                
                # Afficher les notes
                self.afficher_notes()
                messagebox.showinfo("Succès", "Les notes ont été importées avec succès !")

def extraire_notes_from_txt(texte):
    notes = {}
    
    # Motif pour trouver les lignes avec des notes
    # Format: nombre (coeff nombre)
    pattern = r'(\d+[.,]\d+)\s*\(coeff\s*(\d+[.,]\d+)\)'
    
    # Trouver la matière actuelle
    matiere_pattern = r'Code ECUE TBFTR(\d{3})'
    
    matiere_courante = None
    
    for ligne in texte.split('\n'):
        # Chercher si c'est une nouvelle matière
        matiere_match = re.search(matiere_pattern, ligne)
        if matiere_match:
            # Convertir R3.301 en R3.01, R3.302 en R3.02, etc.
            num_matiere = matiere_match.group(1)
            matiere_courante = f"R3.{int(num_matiere):02d}"
        
        # Chercher les notes et coefficients
        if matiere_courante:
            notes_matches = re.finditer(pattern, ligne)
            for match in notes_matches:
                note = float(match.group(1).replace(',', '.'))
                coeff = float(match.group(2).replace(',', '.'))
                
                if matiere_courante not in notes:
                    notes[matiere_courante] = []
                notes[matiere_courante].append((note, coeff))
    
    # Traiter la SAE séparément (format différent)
    sae_pattern = r'TBFTE301.*?(\d+[.,]\d+)\s*\(coeff\s*(\d+[.,]\d+)\)'
    sae_match = re.search(sae_pattern, texte, re.DOTALL)
    if sae_match:
        note = float(sae_match.group(1).replace(',', '.'))
        coeff = float(sae_match.group(2).replace(',', '.'))
        notes['SAE3.01'] = [(note, coeff)]
    
    return notes

def pdf_to_text(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            return text
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la lecture du PDF : {str(e)}")
        return None

if __name__ == "__main__":
    root = tk.Tk()
    app = ApplicationNotes(root)
    root.mainloop() 