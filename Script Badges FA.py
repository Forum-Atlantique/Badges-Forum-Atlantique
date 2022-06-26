# Importation des modules
from PIL import Image, ImageDraw, ImageFont  # Module pour les images
import xlrd  # module pour ouvrir l'excel

# Ouverture du Fichier Excel
document = xlrd.open_workbook("./data.xls")

# Listing des différents fonds de badges
backgroundChauffeur = Image.open("./backgrounds/chauffeur.png")
backgroundStaff = Image.open("./backgrounds/staff.png")
backgroundPilote = Image.open("./backgrounds/pilote.png")
backgroundEntreprise = Image.open("./backgrounds/entreprise.png")

# colors
dark_blue = (27, 53, 81)
light_blue = (93, 188, 210)
blue = (23, 114, 237)
orange = (249, 174, 12)
dark_orange = (255, 111, 0)
yellow_green = (232, 240, 165)
green = (91, 185, 67)

# Fonts
MainFont = r"./fonts/DINBold.ttf"
SubtitleFont = r"./fonts/DINLight.ttf"
SubtitleColor = dark_blue

# Fonctions utilitaires

def center_text(img, font, text1, text2, fill, text1_is_Prenom):
    """
    It takes an image, a font, two strings, and a fill color, and returns an image
    with the two strings centered on the image
    
    Args:
      img: The image to draw on
      font: The font to use.
      text1: The first line of text to be drawn on the image.
      text2: The text to be displayed on the image.
      fill: The color of the text.
    
    Returns:
      The image with the text on it.
    """
    draw = ImageDraw.Draw(img)  # Initialize drawing on the image
    w, h = img.size  # get width and height of image
    t1_width, t1_height = draw.textsize(text1, font)  # Get text1 size
    t2_width, t2_height = draw.textsize(text2, font)
    if text1_is_Prenom:
        p1 = ((w-t1_width)/2, h // 3)  # H-center align text1
    else:
        p1 = ((w-t1_width)/2, h // 3 + h//3)
    p2 = ((w-t2_width)/2, h // 3 + h // 7)  # H-center align text2
    draw.text(p1, text1, fill, font)  # draw text on top of image
    draw.text(p2, text2, fill, font)
    return img


def add_text(img, color, text1, text2, font, font_size, is_staff, text1_is_Prenom):
    """
    It takes an image, a color, two strings, a font, and a font size, and returns an
    image with the two strings centered on the image
    
    Args:
      img: the image to add text to
      color: The color of the text.
      text1: The first line of text to be added to the image.
      text2: The text to be displayed on the image.
      font: the font file to use
      font_size: The size of the font.
    
    Returns:
      The image with the text added.
    """
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(font, size=font_size)
    if is_staff:
        w, h = img.size  # get width and height of image
        t1_width, t1_height = draw.textsize(text1, font)  # Get text1 size
        t2_width, t2_height = draw.textsize(text2, font)
        if text1_is_Prenom:
            # Plus le deuxième terme est grand, plus le texte est bas
            text1_offset = ((w-t1_width)/2, h // 4)
        else:
            text1_offset = p1 = ((w-t1_width)/2, h // 4 + h//3)
        text2_offset = ((w-t2_width)/2, h // 4 + h // 7)
        draw.text(text1_offset, text1, color, font)
        draw.text(text2_offset, text2, color, font)
    else:
        center_text(img, font, text1, text2, color, text1_is_Prenom)
    return img


def write_image(background, color1, color2, Prenom, Nom, text3, is_staff):
    """
    It takes in a background image, two colors, a first name, a last name, and a
    subtitle, and returns an image with the first name and last name in the first
    color and the subtitle in the second color
    
    Args:
      background: the image to write on
      color1: The color of the first line of text
      color2: The color of the text
      Prenom: First name
      Nom: Last name
      text3: The text that will be displayed on the image
    
    Returns:
      The background image with the text added to it.
    """
    add_text(background, color1, Prenom, Nom, MainFont, 45, is_staff, True)
    add_text(background, color2, text3, '', SubtitleFont, 25, is_staff, False)
    return background

# Fonction principale

TYPES_SETTINGS = {
    "chauffeur": {
        "sheet_index": 0
    },
    "pilote": {
        "sheet_index": 1
    },
    "staff": {
        "sheet_index": 2
    },
    "entreprise": {
        "sheet_index": 3
    }
}

def main(type):
    settings = TYPES_SETTINGS[type]
    feuille_i = document.sheet_by_index(settings["sheet_index"])
    rows = feuille_i.nrows
    color1 = orange
    background = backgroundStaff
    for r in range(1, rows):
        Prenom = feuille_i.cell_value(rowx=r, colx=0)
        Nom = feuille_i.cell_value(rowx=r, colx=1)
        text3 = feuille_i.cell_value(rowx=r, colx=2)
        img_name = 'Badge' + Nom + Prenom + '.png'
        is_staff = type == "staff"
        background = write_image(
            background, color1, SubtitleColor, Prenom, Nom, text3, is_staff)
        background.save(img_name)
        background = Image.open(r"./backgrounds/staff.png")

# Commande lors de l'exécution du programme
if __name__ == '__main__':
    print('Vous avez plusieurs possibilités de commande : \n 1 : Création de badges pour les chauffeurs uniquement \n 2 : Création de badges pour les pilotes uniquement \n 3 : Création de badges pour l’équipe organisatrice uniquement \n 4 : Création de badges pour les entreprises uniquement \n 5 : Création de badges pour tout le monde \n 6 : Création de badge pour une seule personne')
    Commande = input("Entrez votre commande: ")
    print(Commande)
    if Commande == '1':  # Chauffeur
        feuille_i = document.sheet_by_index(0)
        rows = feuille_i.nrows
        text3 = ''
        color1 = dark_orange
        background = backgroundChauffeur
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            img_name = 'Badge' + Nom + Prenom + '.png'
            background = write_image(
                background, color1, SubtitleColor, Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"./backgroundChauffeur.png")
    if Commande == '2':  # Pilote
        feuille_i = document.sheet_by_index(1)
        rows = feuille_i.nrows
        color1 = green
        background = backgroundPilote
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom + '.png'
            background = write_image(background, color1, SubtitleColor, Prenom,
                                     Nom, text3)
            background.save(img_name)
            background = Image.open(r"./backgroundPilote.png")
    if Commande == '3':  # Equipe Organisatrice
        main("staff")
    if Commande == '4':  # Entreprise
        feuille_i = document.sheet_by_index(3)
        rows = feuille_i.nrows
        background = backgroundEntreprise
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom + '.png'
            color1 = dark_blue
            background = write_image(
                background, color1, SubtitleColor, Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"./backgroundEntreprise.png")
    if Commande == '5':  # Tout l'excel
        feuille_i = document.sheet_by_index(0)  # CHauffeur
        rows = feuille_i.nrows
        text3 = ''
        color1 = dark_orange
        background = backgroundChauffeur
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            img_name = 'Badge' + Nom + Prenom + '.png'
            background = write_image(
                background, color1, SubtitleColor, Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"./backgroundChauffeur.png")
        feuille_i = document.sheet_by_index(1)  # Pilote
        rows = feuille_i.nrows
        color1 = green
        background = backgroundPilote
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom + '.png'
            background = write_image(background, color1, SubtitleColor, Prenom,
                                     Nom, text3)
            background.save(img_name)
            background = Image.open(r"./backgroundPilote.png")
        feuille_i = document.sheet_by_index(2)  # Equipe Organisatrice
        rows = feuille_i.nrows
        color1 = orange
        background = backgroundStaff
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom + '.png'
            background = write_image(
                background, color1, SubtitleColor, Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"./backgroundStaff.png")
        feuille_i = document.sheet_by_index(3)  # Entreprise
        rows = feuille_i.nrows
        background = backgroundEntreprise
        for r in range(1, rows):
            Prenom = feuille_i.cell_value(rowx=r, colx=0)
            Nom = feuille_i.cell_value(rowx=r, colx=1)
            text3 = feuille_i.cell_value(rowx=r, colx=2)
            img_name = 'Badge' + Nom + Prenom + '.png'
            color1 = dark_blue
            background = write_image(
                background, color1, SubtitleColor, Prenom, Nom, text3)
            background.save(img_name)
            background = Image.open(r"./backgroundEntreprise.png")
    if Commande == '6':  # Un seul nom
        Prenom = input('Prénom: ')
        Nom = input('Nom en Majuscule:')
        img_name = 'Badge' + Nom + Prenom + '.png'
        print('Vous avez plusieurs possibilités de commande : \n 1 : Chauffeur \n 2 : Pilote \n 3 : Equipe Organisatrice \n 4 : Entreprise')
        # renvoie une chaîne de carctère
        Role = input('Entrez votre commande : ')
        print(Role)
        if Role == '1':  # Chauffeur
            text3 = ''
            color1 = dark_orange
            background = backgroundChauffeur
        if Role == '2':  # Pilote
            text3 = ''
            color1 = green
            background = backgroundPilote
        if Role == '3':  # Equipe Organisatrice
            text3 = input('Entrez le rôle: ')
            color1 = orange
            background = backgroundStaff
        if Role == '4':  # Entreprise
            text3 = input('Entrez le nom de l’entreprise : ')
            color1 = dark_blue
            background = backgroundEntreprise
        background = write_image(
            background, color1, SubtitleColor, Prenom, Nom, text3)
        background.save(img_name)
