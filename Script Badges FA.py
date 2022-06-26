## Importation des modules
from PIL import Image, ImageDraw, ImageFont #Module pour les images
import xlrd #module pour ouvrir l'excel

## Ouverture du Fichier Excel
document = xlrd.open_workbook(r"C:\Users\User\Downloads\ScriptFABadges\ExcelBadgesFA.xlsx")

##Listing des différents fonds de badges
backgroundChauffeur = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundChauffeur.png")
backgroundEquipeOrganisatrice = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundEquipeOrganisatrice.png")
backgroundPilote = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundPilote.png")
backgroundEntreprise = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundEntreprise.png")

##colors
dark_blue = (27,53,81)
light_blue = (93,188,210)
blue = (23,114,237)
orange = (249,174,12)
dark_orange = (255,111,0)
yellow_green = (232,240,165)
green = (91, 185, 67)

##Font
MainFont = r"C:\Users\User\Downloads\ScriptFABadges\DINBold.ttf"
SubtitleFont = r"C:\Users\User\Downloads\ScriptFABadges\DINLight.ttf"
SubtitleColor = dark_blue

## Fonctions de centrage et d'ajout de texte à l'image
def center_text(img,font,text1,text2,fill):
    draw = ImageDraw.Draw(img) # Initialize drawing on the image
    w,h = img.size # get width and height of image
    t1_width, t1_height = draw.textsize(text1, font) # Get text1 size
    t2_width, t2_height = draw.textsize(text2, font)
    if text1 == Prenom :
        p1 = ((w-t1_width)/2,h // 3) # H-center align text1
    else :
        p1 = ((w-t1_width)/2,h // 3 + h//3)
    p2 = ((w-t2_width)/2,h // 3 + h // 7) # H-center align text2
    draw.text(p1, text1, fill, font) # draw text on top of image
    draw.text(p2, text2, fill, font)
    return img

def add_text(img,color,text1,text2,font,font_size):
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(font,size=font_size)
    if Commande == '3' or Role == '3' :
        w,h = img.size # get width and height of image
        t1_width, t1_height = draw.textsize(text1, font) # Get text1 size
        t2_width, t2_height = draw.textsize(text2, font)
        if text1 == Prenom :
            text1_offset = ((w-t1_width)/2,h //4)#Plus le deuxième terme est grand, plus le texte est bas
        else :
            text1_offset = p1 = ((w-t1_width)/2,h // 4 + h//3)
        text2_offset = ((w-t2_width)/2,h // 4 + h // 7)
        draw.text(text1_offset, text1, color, font)
        draw.text(text2_offset, text2, color, font)
    else :
        center_text(img,font,text1,text2, color)
    return img


def write_image(background,color1,color2,Prenom,Nom,text3):
    add_text(background,color1,Prenom,Nom,MainFont, 45)
    add_text(background,color2,text3, '',SubtitleFont, 25)
    return background

## Commande lors de l'exécution du programme
if __name__ == '__main__':
    print ('Vous avez plusieurs possibilités de commande : \n 1 : Création de badges pour les chauffeurs uniquement \n 2 : Création de badges pour les pilotes uniquement \n 3 : Création de badges pour l’équipe organisatrice uniquement \n 4 : Création de badges pour les entreprises uniquement \n 5 : Création de badges pour tout le monde \n 6 : Création de badge pour une seule personne')
    Commande = input("Entrez votre commande: ")
    print (Commande)
    if Commande == '1':# Chauffeur
        feuille_i = document.sheet_by_index(0)
        rows = feuille_i.nrows
        text3 = ''
        color1 = dark_orange
        background = backgroundChauffeur
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            img_name = 'Badge' + Nom + Prenom +  '.png'
            background = write_image(background,color1,SubtitleColor, Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundChauffeur.png")
    if Commande == '2' :#Pilote
        feuille_i = document.sheet_by_index(1)
        rows = feuille_i.nrows
        color1 = green
        background = backgroundPilote
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom +  '.png'
            background = write_image(background,color1,SubtitleColor, Prenom,
Nom, text3)
            background.save(img_name)
            background = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundPilote.png")
    if Commande == '3' :#Equipe Organisatrice
        feuille_i = document.sheet_by_index(2)
        rows = feuille_i.nrows
        color1 = orange
        background = backgroundEquipeOrganisatrice
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom +  '.png'
            background = write_image(background,color1,SubtitleColor,Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundEquipeOrganisatrice.png")
    if Commande == '4' :#Entreprise
        feuille_i = document.sheet_by_index(3)
        rows = feuille_i.nrows
        background = backgroundEntreprise
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom +  '.png'
            color1 = dark_blue
            background = write_image(background,color1,SubtitleColor,Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundEntreprise.png")
    if Commande == '5' :  # Tout l'excel
        feuille_i = document.sheet_by_index(0)#CHauffeur
        rows = feuille_i.nrows
        text3 = ''
        color1 = dark_orange
        background = backgroundChauffeur
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            img_name = 'Badge' + Nom + Prenom +  '.png'
            background = write_image(background,color1,SubtitleColor, Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundChauffeur.png")
        feuille_i = document.sheet_by_index(1)#Pilote
        rows = feuille_i.nrows
        color1 = green
        background = backgroundPilote
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom +  '.png'
            background = write_image(background,color1,SubtitleColor, Prenom,
Nom, text3)
            background.save(img_name)
            background = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundPilote.png")
        feuille_i = document.sheet_by_index(2)#Equipe Organisatrice
        rows = feuille_i.nrows
        color1 = orange
        background = backgroundEquipeOrganisatrice
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom +  '.png'
            background = write_image(background,color1,SubtitleColor,Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundEquipeOrganisatrice.png")
        feuille_i = document.sheet_by_index(3)#Entreprise
        rows = feuille_i.nrows
        background = backgroundEntreprise
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom +  '.png'
            color1 = dark_blue
            background = write_image(background,color1,SubtitleColor,Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"C:\Users\User\Downloads\ScriptFABadges\backgroundEntreprise.png")
    if Commande == '6' :#Un seul nom
        Prenom = input('Prénom: ')
        Nom = input('Nom en Majuscule:')
        img_name = 'Badge' + Nom + Prenom +  '.png'
        print ('Vous avez plusieurs possibilités de commande : \n 1 : Chauffeur \n 2 : Pilote \n 3 : Equipe Organisatrice \n 4 : Entreprise')
        Role = input('Entrez votre commande : ')#renvoie une chaîne de carctère
        print (Role)
        if Role == '1' :#Chauffeur
            text3 = ''
            color1 = dark_orange
            background = backgroundChauffeur
        if Role == '2' :#Pilote
            text3 = ''
            color1 = green
            background = backgroundPilote
        if Role == '3' :#Equipe Organisatrice
            text3 = input('Entrez le rôle: ')
            color1 = orange
            background = backgroundEquipeOrganisatrice
        if Role == '4' :#Entreprise
            text3 = input('Entrez le nom de l’entreprise : ')
            color1 = dark_blue
            background = backgroundEntreprise
        background = write_image(background,color1,SubtitleColor,Prenom, Nom, text3)
        background.save(img_name)

