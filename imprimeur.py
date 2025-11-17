import os
import sys
import subprocess
import winreg
import win32api
import win32print
import tkinter as tk
from tkinter import messagebox, Listbox, Button, END

def installer_dependances():
    """Vérifie et installe automatiquement PyPDF2 et pywin32 si nécessaire."""
    try:
        import PyPDF2
        import win32api
    except ImportError:
        messagebox.showinfo("Information", "Installation des dépendances en cours...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "PyPDF2", "pywin32"])
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'installer les dépendances : {e}")
            sys.exit(1)

def lister_imprimantes():
    """Liste les imprimantes installées sur le système."""
    try:
        imprimantes = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
        return [imprimante[2] for imprimante in imprimantes]
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible de lister les imprimantes : {e}")
        return []

def ajouter_imprimante_menu_contextuel(nom_imprimante):
    """Ajoute une option au menu contextuel des fichiers PDF pour imprimer avec l'imprimante sélectionnée."""
    try:
        pdf_key = r"Software\Classes\SystemFileAssociations\.pdf\shell\ImprimerPDF"
        command_key = r"Software\Classes\SystemFileAssociations\.pdf\shell\ImprimerPDF\command"

        script_path = os.path.abspath(__file__)

        # --------------------------
        # FIX : commande différente si .exe
        # --------------------------
        if getattr(sys, 'frozen', False):
            # On tourne en version compilée (PyInstaller)
            commande = f'"{sys.executable}" "%1" "{nom_imprimante}"'
        else:
            # Version python normale
            commande = f'"{sys.executable}" "{script_path}" "%1" "{nom_imprimante}"'

        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, pdf_key) as key:
            winreg.SetValue(key, "", winreg.REG_SZ, f"Imprimer avec {nom_imprimante}")

        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, command_key) as key:
            winreg.SetValue(key, "", winreg.REG_SZ, commande)

        messagebox.showinfo("Succès", f"L'option 'Imprimer avec {nom_imprimante}' a été ajoutée au menu contextuel des PDF.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

def supprimer_cle_recursive(cle_parent, sous_cle):
    """Supprime récursivement une clé de registre et toutes ses sous-clés."""
    try:
        cle = winreg.OpenKey(cle_parent, sous_cle, 0, winreg.KEY_ALL_ACCESS)
        
        try:
            i = 0
            while True:
                nom_sous_cle = winreg.EnumKey(cle, i)
                supprimer_cle_recursive(cle, nom_sous_cle)
        except OSError:
            pass
        
        winreg.CloseKey(cle)
        winreg.DeleteKey(cle_parent, sous_cle)
    except FileNotFoundError:
        pass
    except Exception as e:
        raise e

def supprimer_imprimante_menu_contextuel():
    """Supprime complètement l'option du menu contextuel des fichiers PDF."""
    try:
        base_key = r"Software\Classes\SystemFileAssociations\.pdf\shell"
        
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_key, 0, winreg.KEY_ALL_ACCESS) as parent_key:
            supprimer_cle_recursive(parent_key, "ImprimerPDF")
        
        messagebox.showinfo("Succès", "L'option a été supprimée du menu contextuel des PDF.")
    except FileNotFoundError:
        messagebox.showinfo("Information", "L'option n'existe pas dans le menu contextuel.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la suppression : {e}")

def imprimer_pdf(fichier_pdf, nom_imprimante):
    """Imprime un fichier PDF sans l'ouvrir en utilisant win32print."""
    try:
        h_printer = win32print.OpenPrinter(nom_imprimante)
        try:
            with open(fichier_pdf, "rb") as file:
                pdf_data = file.read()

            job_info = {"pDatatype": "RAW", "pDevMode": None, "pSecurity": None}
            win32print.StartDocPrinter(h_printer, 1, (fichier_pdf, None, "RAW"))
            win32print.StartPagePrinter(h_printer)
            win32print.WritePrinter(h_printer, pdf_data)
            win32print.EndPagePrinter(h_printer)
            win32print.EndDocPrinter(h_printer)
            messagebox.showinfo("Succès", f"Le fichier {fichier_pdf} a été envoyé à l'imprimante {nom_imprimante}.")
        finally:
            win32print.ClosePrinter(h_printer)
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de l'impression : {e}")

def interface_graphique():
    """Affiche une interface pour gérer les imprimantes."""
    root = tk.Tk()
    root.title("Gestion des Imprimantes pour PDF")

    imprimantes = lister_imprimantes()
    if not imprimantes:
        messagebox.showerror("Erreur", "Aucune imprimante trouvée.")
        return

    listbox = Listbox(root, width=50, height=10)
    for imprimante in imprimantes:
        listbox.insert(END, imprimante)
    listbox.pack(pady=10)

    def ajouter_imprimante():
        try:
            selection = listbox.get(listbox.curselection())
            ajouter_imprimante_menu_contextuel(selection)
        except:
            messagebox.showwarning("Attention", "Veuillez sélectionner une imprimante.")

    def supprimer_imprimante():
        if messagebox.askyesno("Confirmation", "Voulez-vous vraiment supprimer l'option du menu contextuel ?"):
            supprimer_imprimante_menu_contextuel()

    bouton_ajouter = Button(root, text="Ajouter au menu contextuel", command=ajouter_imprimante)
    bouton_ajouter.pack(pady=5)

    bouton_supprimer = Button(root, text="Supprimer du menu contextuel", command=supprimer_imprimante)
    bouton_supprimer.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    installer_dependances()

    # Mode impression (depuis menu contextuel)
    if len(sys.argv) == 3:
        fichier_pdf = sys.argv[1]
        nom_imprimante = sys.argv[2]
        imprimer_pdf(fichier_pdf, nom_imprimante)

    # Mode interface graphique
    else:
        interface_graphique()
