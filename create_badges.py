# Importation des modules
from PIL import Image, ImageDraw, ImageFont  # Module pour les images
import xlrd  # module pour ouvrir l'excel

# colors
dark_blue = (27, 53, 81)
light_blue = (93, 188, 210)
blue = (23, 114, 237)
orange = (249, 174, 12)
dark_orange = (255, 111, 0)
yellow_green = (232, 240, 165)
green = (91, 185, 67)

# Fonts
TITLE_FONT = r"./fonts/DINBold.ttf"
TEXT_FONT = r"./fonts/DINLight.ttf"
SECONDARY_COLOR = dark_blue

# Fichier Excel avec les noms
DATA_FILE_NAME = "./data.xls"

# Dossier où enregistrer les images créées
OUTPUT_DIR = "./badges"

# Paramètres pour chaque type de badge
TYPES_SETTINGS = {
    "chauffeur": {
        "sheet_index": 0,
        "color": dark_orange,
        "bg": "backgrounds/chauffeur.png",
    },
    "pilote": {
        "sheet_index": 1,
        "color": green,
        "bg": "backgrounds/pilote.png",
    },
    "staff": {
        "sheet_index": 2,
        "color": orange,
        "bg": "backgrounds/staff.png",
    },
    "entreprise": {
        "sheet_index": 3,
        "color": dark_blue,
        "bg": "backgrounds/entreprise.png",
    }
}


# Fonctions utilitaires

def center_text(img, font, text1, text2, fill, test1_is_name):
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
    if test1_is_name:
        p1 = ((w-t1_width)/2, h // 3)  # H-center align text1
    else:
        p1 = ((w-t1_width)/2, h // 3 + h//3)
    p2 = ((w-t2_width)/2, h // 3 + h // 7)  # H-center align text2
    draw.text(p1, text1, fill, font)  # draw text on top of image
    draw.text(p2, text2, fill, font)
    return img


def add_text(img, color, text1, text2, font, font_size, is_staff, test1_is_name):
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
        if test1_is_name:
            # Plus le deuxième terme est grand, plus le texte est bas
            text1_offset = ((w-t1_width)/2, h // 4)
        else:
            text1_offset = p1 = ((w-t1_width)/2, h // 4 + h//3)
        text2_offset = ((w-t2_width)/2, h // 4 + h // 7)
        draw.text(text1_offset, text1, color, font)
        draw.text(text2_offset, text2, color, font)
    else:
        center_text(img, font, text1, text2, color, test1_is_name)
    return img


def write_image(background, color1, color2, first_name, last_name, role, is_staff):
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
    add_text(background, color1, first_name, last_name, TITLE_FONT, 45, is_staff, True)
    add_text(background, color2, role, '', TEXT_FONT, 25, is_staff, False)
    return background


# Fonction principale
def main(type):
    """Open the Excel file, reads the data, and generates the badges
    
    Args:
      type: The type of badge you want to create.
    """
    settings = TYPES_SETTINGS[type]
    document = xlrd.open_workbook(DATA_FILE_NAME)
    sheet = document.sheet_by_index(settings["sheet_index"])
    nb_rows = sheet.nrows
    primary_color = settings["color"]
    background_file_name = settings["bg"]
    for r in range(1, nb_rows):
        first_name = sheet.cell_value(rowx=r, colx=0)
        last_name = sheet.cell_value(rowx=r, colx=1)
        role = sheet.cell_value(rowx=r, colx=2)
        image = Image.open(background_file_name)
        image = write_image(
            image,
            primary_color,
            SECONDARY_COLOR,
            first_name,
            last_name,
            role,
            type == "staff"
        )
        file_name = f'{OUTPUT_DIR}/Badge {first_name} {last_name}.png'
        image.save(file_name)


# Commande lors de l'exécution du programme
if __name__ == '__main__':
    print('Vous avez plusieurs possibilités de commande : \n 1 : Création de badges pour les chauffeurs uniquement \n 2 : Création de badges pour les pilotes uniquement \n 3 : Création de badges pour l’équipe organisatrice uniquement \n 4 : Création de badges pour les entreprises uniquement \n 5 : Création de badges pour tout le monde \n 6 : Création de badge pour une seule personne')
    input_cmd = input("Entrez votre commande: ")
    if input_cmd == '1': main("chauffeur")
    elif input_cmd == '2': main("pilote")
    elif input_cmd == '3': main("staff")
    elif input_cmd == '4': main("entreprise")
    elif input_cmd == '5':  # Tout l'excel
        for type in TYPES_SETTINGS:
            main(type)
    elif input_cmd == '6':  # Un seul nom
        first_name = input('Prénom: ')
        last_name = input('Nom en Majuscule: ')
        print('Vous avez plusieurs possibilités de commande : \n 1 : Chauffeur \n 2 : Pilote \n 3 : Equipe Organisatrice \n 4 : Entreprise')
        # lire le choix
        input_type = input('Entrez votre commande : ')
        if input_type == '1': type = "chauffeur"
        elif input_type == '2': type = "pilote"
        elif input_type == "3": type = "staff"
        elif input_type == '4': type = "entreprise"
        else: raise Exception("Choix invalide !")
        # définir les paramètres
        settings = TYPES_SETTINGS[type]
        primary_color = settings["color"]
        background_file_name = settings["bg"]
        role = ""
        if input_type == '3':  # Equipe Organisatrice
            role = input('Entrez le rôle: ')
        if input_type == '4':  # Entreprise
            text3 = input('Entrez le nom de l’entreprise : ')
        # créer l'image
        image = Image.open(background_file_name)
        image = write_image(
            image,
            primary_color,
            SECONDARY_COLOR,
            first_name,
            last_name,
            role,
            type == "staff"
        )
        file_name = f'Badge {first_name} {last_name}.png'
        image.save(file_name)
    else:
        raise Exception("Choix invalide ! Veuillez entrez un nombre entre 1 et 6.")
