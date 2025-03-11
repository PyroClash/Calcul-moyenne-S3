import sys
import subprocess
import importlib.util

def check_and_install_packages():
    required_packages = ['tkinter', 'PyPDF2', 'pandas']
    
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
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PyPDF2 import PdfReader
import pandas as pd
from datetime import datetime
# Importer les matières et les coefficients depuis le nouveau fichier
from matieres_coeffs import MATIERES, COEFFICIENTS

class ApplicationNotes:
    def __init__(self, root):
        self.root = root
        self.root.title("Calcul des moyennes UE - Made by PDO")
        
        # Dictionnaire pour stocker les notes de chaque matière
        self.notes_par_matiere = {}
        # Résultats pour le mode batch
        self.resultats_batch = []
        
        # Frame principal
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame pour le chargement PDF
        self.pdf_frame = ttk.LabelFrame(self.main_frame, text="Chargement des notes", padding="5")
        self.pdf_frame.pack(fill=tk.X, pady=5)
        
        # Checkbox pour activer le mode batch
        self.mode_batch = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.pdf_frame, text="Mode batch (plusieurs relevés)", 
                        variable=self.mode_batch).pack(side=tk.LEFT, padx=5, pady=10)
        
        # Bouton pour charger un relevé PDF
        ttk.Button(self.pdf_frame, text="Charger relevé(s) PDF", 
                  command=self.charger_pdf).pack(side=tk.LEFT, padx=5, pady=10)
        
        # Checkbox pour ignorer SAE3.01
        self.ignore_sae = tk.BooleanVar(value=True)  # Activée par défaut
        ttk.Checkbutton(self.pdf_frame, text="Ignorer et neutraliser SAE3.01", 
                        variable=self.ignore_sae).pack(side=tk.RIGHT, padx=5, pady=10)
        
        # Barre de progression (cachée par défaut)
        self.progress_frame = ttk.Frame(self.main_frame)
        self.progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, 
                                            mode='determinate', length=400)
        self.progress_bar.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)
        
        self.progress_label = ttk.Label(self.progress_frame, text="")
        self.progress_label.pack(side=tk.RIGHT, padx=5)
        
        # Cache la barre de progression au début
        self.progress_frame.pack_forget()
        
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
            self.resultats_batch = []
    
    def calculer_moyennes(self, pdf_path=None):
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
        
        # Si on est en mode batch, on ne réinitialise pas les résultats
        if not self.mode_batch.get() or pdf_path is None:
            self.resultats.delete(1.0, tk.END)
            self.resultats.insert(tk.END, "Moyennes par UE :\n")
            self.resultats.insert(tk.END, "-" * 40 + "\n")
        
        # Dictionnaire pour stocker les moyennes par UE et par ressource
        moyennes_ue = {}
        moyennes_ressources = {}
        
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
                        moyennes_ressources[matiere] = moyenne_matiere
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
                moyennes_ue[ue] = moyenne
                
                if not self.mode_batch.get() or pdf_path is None:
                    self.resultats.insert(tk.END, f"{ue}: {moyenne:.2f}/20\n")
                print(f"  Moyenne {ue}: {moyenne:.2f}/20")
            else:
                moyennes_ue[ue] = None
                
                if not self.mode_batch.get() or pdf_path is None:
                    self.resultats.insert(tk.END, f"{ue}: Notes manquantes\n")
                print(f"  {ue}: Notes manquantes")
        
        # Calculer la moyenne générale (moyenne des moyennes d'UE)
        moyennes_valides = [moy for moy in moyennes_ue.values() if moy is not None]
        if moyennes_valides:
            moyenne_generale = sum(moyennes_valides) / len(moyennes_valides)
            
            if not self.mode_batch.get() or pdf_path is None:
                self.resultats.insert(tk.END, "-" * 40 + "\n")
                self.resultats.insert(tk.END, f"Moyenne générale: {moyenne_generale:.2f}/20\n")
            print(f"  Moyenne générale: {moyenne_generale:.2f}/20")
        else:
            moyenne_generale = None
            
            if not self.mode_batch.get() or pdf_path is None:
                self.resultats.insert(tk.END, "-" * 40 + "\n")
                self.resultats.insert(tk.END, f"Moyenne générale: Notes insuffisantes\n")
            print(f"  Moyenne générale: Notes insuffisantes")
        
        # Si on est en mode batch, on stocke les résultats pour le CSV
        if self.mode_batch.get() and pdf_path:
            # Extraire le nom et prénom du fichier
            nom_prenom = self.extraire_nom_prenom(pdf_path)
            
            resultat_etudiant = {
                'Nom': nom_prenom['nom'],
                'Prénom': nom_prenom['prenom'],
                'Moyenne générale': moyenne_generale
            }
            
            # Ajouter les moyennes par UE
            for ue, moyenne in moyennes_ue.items():
                if moyenne is not None:
                    resultat_etudiant[ue] = moyenne
                else:
                    resultat_etudiant[ue] = float('nan')
            
            # Ajouter les moyennes par ressource
            for matiere, moyenne in moyennes_ressources.items():
                resultat_etudiant[matiere] = moyenne
            
            self.resultats_batch.append(resultat_etudiant)
        
        return moyennes_ue, moyenne_generale
    
    def extraire_nom_prenom(self, pdf_path):
        # Extraire le nom du fichier sans le chemin
        nom_fichier = os.path.basename(pdf_path)
        # Format attendu: Releve-NOM-PRENOM-TBFS3T-2024-2025.pdf
        match = re.match(r'Releve-([^-]+)-([^-]+)-TBFS3T-\d{4}-\d{4}', nom_fichier)
        
        if match:
            return {'nom': match.group(1), 'prenom': match.group(2)}
        else:
            # Si le format ne correspond pas, utiliser le nom du fichier
            return {'nom': nom_fichier, 'prenom': ''}

    def charger_pdf(self):
        if self.mode_batch.get():
            # Mode batch: sélection de plusieurs fichiers
            pdf_paths = filedialog.askopenfilenames(
                title="Sélectionner les relevés de notes PDF",
                filetypes=[("Fichiers PDF", "*.pdf")]
            )
            
            if pdf_paths:
                # Réinitialiser les résultats batch
                self.resultats_batch = []
                
                # Afficher la barre de progression
                self.progress_frame.pack(fill=tk.X, pady=5, after=self.pdf_frame)
                self.progress_var.set(0)
                self.progress_label.config(text="0%")
                
                # Traiter chaque fichier
                total_files = len(pdf_paths)
                for i, pdf_path in enumerate(pdf_paths):
                    # Mettre à jour la barre de progression
                    progress_pct = (i / total_files) * 100
                    self.progress_var.set(progress_pct)
                    self.progress_label.config(text=f"{int(progress_pct)}%")
                    self.root.update_idletasks()
                    
                    # Convertir le PDF en texte
                    texte = pdf_to_text(pdf_path)
                    if texte:
                        # Extraire les notes
                        self.notes_par_matiere = extraire_notes_from_txt(texte)
                        
                        # Afficher les notes du fichier courant
                        self.afficher_notes()
                        
                        # Calculer les moyennes pour ce relevé
                        self.calculer_moyennes(pdf_path)
                
                # Terminer la barre de progression
                self.progress_var.set(100)
                self.progress_label.config(text="100%")
                
                # Générer le CSV avec les résultats
                if self.resultats_batch:
                    self.generer_csv()
                
                # Cacher la barre de progression après quelques secondes
                self.root.after(2000, lambda: self.progress_frame.pack_forget())
        else:
            # Mode normal: sélection d'un seul fichier
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
    
    def generer_csv(self):
        # Convertir la liste de dictionnaires en DataFrame
        df = pd.DataFrame(self.resultats_batch)
        
        # Trier par moyenne générale décroissante
        df = df.sort_values(by='Moyenne générale', ascending=False)
        
        # Ajouter le rang
        df['Rang'] = range(1, len(df) + 1)
        
        # Réorganiser les colonnes (nom, prénom, rang, moyenne générale, puis le reste)
        colonnes = ['Nom', 'Prénom', 'Rang', 'Moyenne générale'] + [
            col for col in df.columns if col not in ['Nom', 'Prénom', 'Rang', 'Moyenne générale']
        ]
        df = df[colonnes]
        
        # Définir le nom du fichier avec la date
        date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Fichiers CSV", "*.csv")],
            initialfile=f"resultats_promotion_{date_str}.csv"
        )
        
        if file_path:
            # Sauvegarder en CSV
            df.to_csv(file_path, index=False, sep=';', decimal=',')
            messagebox.showinfo("Succès", f"Fichier CSV généré avec succès !\n{file_path}")
            
            # Afficher un résumé dans la zone de résultats
            self.resultats.delete(1.0, tk.END)
            self.resultats.insert(tk.END, f"Résumé du traitement batch ({len(df)} étudiants) :\n")
            self.resultats.insert(tk.END, "-" * 40 + "\n")
            
            # Afficher les 3 premiers et les 3 derniers
            top3 = df.head(3)
            bottom3 = df.tail(3)
            
            self.resultats.insert(tk.END, "Top 3 :\n")
            for _, row in top3.iterrows():
                self.resultats.insert(tk.END, f"{row['Rang']}. {row['Prénom']} {row['Nom']}: {row['Moyenne générale']:.2f}/20\n")
            
            self.resultats.insert(tk.END, "\nDerniers :\n")
            for _, row in bottom3.iterrows():
                self.resultats.insert(tk.END, f"{row['Rang']}. {row['Prénom']} {row['Nom']}: {row['Moyenne générale']:.2f}/20\n")
            
            self.resultats.insert(tk.END, f"\nMoyenne de la promotion: {df['Moyenne générale'].mean():.2f}/20\n")

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