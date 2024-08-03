from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
import math
import tkinter as tk
from tkinter import simpledialog

# Prix des accessoires (à ajuster selon les prix réels)
PRIX_CLAMP_FIN = 7.50
PRIX_CLAMP_MILIEU = 7.50
PRIX_RAIL_PAR_METRE = 30.83  
PRIX_CONNEXION = 15
PRIX_ECROU_BOULON_2_60 = 1.5
PRIX_ECROU_BOULON_2_80 = 1.5
PRIX_BETON = 120  
LONGUEUR_RAIL = 6

# Marge d'erreur
MARGE_ERREUR = 0.0

# Taux de TVA
TAUX_TVA = 0.2  # 20% de TVA

# Fonction pour appliquer la marge d'erreur
def appliquer_marge_erreur(valeur):
    return math.ceil(valeur + (valeur * MARGE_ERREUR))

# Fonction pour formater les montants en type monétaire
def format_monetaire(montant):
    return f"{montant:.2f} DH"

# Fonction pour calculer les accessoires
def calcul_accessoires(nombre_panneaux, longueur_panneau, largeur_panneau):
    # Calcul des clamps avec tolérance
    clamps_fin = 4  # 4 clamps de fin nécessaires (2 par extrémité)
    clamps_milieu = (nombre_panneaux - 1) * 2  # 2 clamps de milieu par connexion entre panneaux
    
    # Calcul du nombre de pieds
    nombre_de_pieds = ((nombre_panneaux * largeur_panneau) / 1.5)

    # Calcul des rails horizontaux
    longueur_totale_rails_horizontal = nombre_panneaux * largeur_panneau * 2
    longueur_rail_de_pieds = nombre_de_pieds * 0.3
    longueur_totale_rails_vertical = longueur_panneau * nombre_de_pieds
    
    # Calcul des connexions jumelages
    connexions_jumelage = (3 * nombre_de_pieds)
    boulon_2_60 = 8 * nombre_de_pieds
    boulon_2_80 = 2 * nombre_de_pieds
    
    # Calcul des fondations en béton
    nombre_betons = int(2 * ((nombre_panneaux * largeur_panneau) / 1.5))
    
    # Calcul de la longueur totale des rails
    longueur_totale_rails = longueur_totale_rails_horizontal + longueur_totale_rails_vertical + longueur_rail_de_pieds

    # Appliquer la marge d'erreur
    longueur_totale_rails = round(appliquer_marge_erreur(longueur_totale_rails), 2)

    # Calcul du nombre de rails de 6 mètres
    nombre_rails = math.ceil(longueur_totale_rails / LONGUEUR_RAIL)

    # Appliquer la marge d'erreur et arrondir à l'entier supérieur si nécessaire
    clamps_fin = appliquer_marge_erreur(clamps_fin)
    clamps_milieu = appliquer_marge_erreur(clamps_milieu)
    connexions_jumelage = appliquer_marge_erreur(connexions_jumelage)
    nombre_betons = appliquer_marge_erreur(nombre_betons)
    
    # Calcul du coût total HT
    cout_clamps_fin_HT = clamps_fin * PRIX_CLAMP_FIN
    cout_clamps_milieu_HT = clamps_milieu * PRIX_CLAMP_MILIEU
    cout_rails_HT = nombre_rails * PRIX_RAIL_PAR_METRE
    cout_connexions_HT = connexions_jumelage * PRIX_CONNEXION
    cout_ecrous_boulons_60_HT = boulon_2_60 * PRIX_ECROU_BOULON_2_60
    cout_ecrous_boulons_80_HT = boulon_2_80 * PRIX_ECROU_BOULON_2_80
    cout_beton_HT = nombre_betons * PRIX_BETON
    
    cout_total_HT = (
        cout_clamps_fin_HT + 
        cout_clamps_milieu_HT +
        cout_rails_HT +
        cout_connexions_HT +
        cout_ecrous_boulons_60_HT +
        cout_ecrous_boulons_80_HT +
        cout_beton_HT
    )
    
    # Calcul du coût total TTC
    cout_total_TTC = cout_total_HT * (1 + TAUX_TVA)

    accessoires = {
        "clamps_fin": clamps_fin,
        "clamps_milieu": clamps_milieu,
        "nombre_rails": nombre_rails,
        "connexions_jumelage": connexions_jumelage,
        "boulon_2_60": boulon_2_60,
        "boulon_2_80": boulon_2_80,
        "cout_clamps_fin_HT": format_monetaire(round(cout_clamps_fin_HT, 2)),
        "cout_clamps_milieu_HT": format_monetaire(round(cout_clamps_milieu_HT, 2)),
        "cout_rails_HT": format_monetaire(round(cout_rails_HT, 2)),
        "cout_connexions_HT": format_monetaire(round(cout_connexions_HT, 2)),
        "cout_ecrous_boulons_60_HT": format_monetaire(round(cout_ecrous_boulons_60_HT, 2)),
        "cout_ecrous_boulons_80_HT": format_monetaire(round(cout_ecrous_boulons_80_HT, 2)),
        "cout_beton_HT": format_monetaire(round(cout_beton_HT, 2)),
        "cout_total_HT": format_monetaire(round(cout_total_HT, 2)),
        "cout_total_TTC": format_monetaire(round(cout_total_TTC, 2))
    }

    return accessoires

# Fonction pour générer le PDF
def generer_pdf(nombre_panneaux, longueur_panneau, largeur_panneau, accessoires):
    nom_fichier = "rapport_accessoires.pdf"
    document = SimpleDocTemplate(nom_fichier, pagesize=A4)
    elements = []
    
    styles = getSampleStyleSheet()
    custom_style = ParagraphStyle(
        'CustomStyle',
        parent=styles['Title'],
        textColor=colors.HexColor('#000080'),
        alignment=TA_CENTER
    )
    titre = Paragraph(f"Rapport d'accessoires pour <font color='red'>{nombre_panneaux}</font> panneaux solaires", custom_style)
    
    # Ajouter le logo en en-tête
    logo = "LOGO1.png"  # Assurez-vous que le fichier LOGO1.png existe
    img_logo = Image(logo, width=100, height=50)  # Dimensions ajustées
    img_logo.hAlign = 'LEFT'
    
    elements.append(img_logo)
    elements.append(Spacer(1, 20))
    
    elements.append(titre)
    elements.append(Spacer(1, 24))

    # Ajouter la phrase au-dessus du tableau
    note = Paragraph("* Merci de vérifier les PRIX des accessoires", styles['Normal'])
    elements.append(note)
    elements.append(Spacer(1, 12))
    
    # Table des accessoires
    data = [
        ["Type d'Accessoire", "Quantité", "Prix Unitaire HT (DH)", "Prix Total HT (DH)", "Prix Total TTC (DH)"],
        ["Clamps de fin", accessoires['clamps_fin'], format_monetaire(PRIX_CLAMP_FIN), accessoires['cout_clamps_fin_HT'], format_monetaire(round(float(accessoires['cout_clamps_fin_HT'].replace(' DH', '')) * (1 + TAUX_TVA), 2))],
        ["Clamps du milieu", accessoires['clamps_milieu'], format_monetaire(PRIX_CLAMP_MILIEU), accessoires['cout_clamps_milieu_HT'], format_monetaire(round(float(accessoires['cout_clamps_milieu_HT'].replace(' DH', '')) * (1 + TAUX_TVA), 2))],
        ["Nombre de rails de 6 mètres", accessoires['nombre_rails'], format_monetaire(PRIX_RAIL_PAR_METRE), accessoires['cout_rails_HT'], format_monetaire(round(float(accessoires['cout_rails_HT'].replace(' DH', '')) * (1 + TAUX_TVA), 2))],
        ["Connexions jumelage", accessoires['connexions_jumelage'], format_monetaire(PRIX_CONNEXION), accessoires['cout_connexions_HT'], format_monetaire(round(float(accessoires['cout_connexions_HT'].replace(' DH', '')) * (1 + TAUX_TVA), 2))],
        ["Boulons & écrous 60 mm", accessoires['boulon_2_60'], format_monetaire(PRIX_ECROU_BOULON_2_60), accessoires['cout_ecrous_boulons_60_HT'], format_monetaire(round(float(accessoires['cout_ecrous_boulons_60_HT'].replace(' DH', '')) * (1 + TAUX_TVA), 2))],
        ["Boulons & écrous 80 mm", accessoires['boulon_2_80'], format_monetaire(PRIX_ECROU_BOULON_2_80), accessoires['cout_ecrous_boulons_80_HT'], format_monetaire(round(float(accessoires['cout_ecrous_boulons_80_HT'].replace(' DH', '')) * (1 + TAUX_TVA), 2))],
        ["Fondations en b\u00E9ton", accessoires['nombre_betons'], format_monetaire(PRIX_BETON), accessoires['cout_beton_HT'], format_monetaire(round(float(accessoires['cout_beton_HT'].replace(' DH', '')) * (1 + TAUX_TVA), 2))],
        ["Coût total", "", "", accessoires['cout_total_HT'], accessoires['cout_total_TTC']]
    ]
    
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)
    
    document.build(elements)
    print(f"Le fichier PDF '{nom_fichier}' a été créé avec succès.")

# Fonction principale
def main():
    root = tk.Tk()
    root.withdraw()

    # Demander les données à l'utilisateur
    nombre_panneaux = simpledialog.askinteger("Nombre de panneaux", "Entrez le nombre de panneaux solaires:")
    longueur_panneau = simpledialog.askfloat("Longueur du panneau (m)", "Entrez la longueur du panneau solaire en mètres:")
    largeur_panneau = simpledialog.askfloat("Largeur du panneau (m)", "Entrez la largeur du panneau solaire en mètres:")

    # Calcul des accessoires
    accessoires = calcul_accessoires(nombre_panneaux, longueur_panneau, largeur_panneau)

    # Générer le PDF
    generer_pdf(nombre_panneaux, longueur_panneau, largeur_panneau, accessoires)

# Exécuter le programme principal
if __name__ == "__main__":
    main()
