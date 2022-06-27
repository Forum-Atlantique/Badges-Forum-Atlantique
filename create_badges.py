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


def center_text(
    img: Image,
    font: str,
    text1: str,
    text2: str,
    fill: str,
    test1_is_name: bool
):
    """
    It takes an image, a font, two strings, and a fill color, and returns an image
    with the two strings centered on the image

    Args:
      img: The image to draw on
      font: The font to use.
      text1: The first line of text to be drawn on the image.
      text2: The text to be displayed on the image.
      fill: The color of the text.
      text1_is_name: Indicates if the text1 variable represents the first name

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


def add_text(
    image: Image,
    color: tuple,
    text1: str,
    text2: str,
    font: str,
    font_size: int,
    is_staff: bool,
    test1_is_name: bool
):
    """
    It takes an image, a color, two strings, a font, and a font size, and returns an
    image with the two strings centered on the image

    Args:
      image: the image to add text to
      color: The color of the text.
      text1: The first line of text to be added to the image.
      text2: The text to be displayed on the image.
      font: the font file to use
      font_size: The size of the font.
      is_staff: Indicates if the image is for staff format
      text1_is_name: Indicates if the text1 variable represents the first name

    Returns:
      The image with the text added.
    """

    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(font, size=font_size)
    if is_staff:
        w, h = image.size  # get width and height of image
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
        center_text(image, font, text1, text2, color, test1_is_name)
    return image


def create_image(
    background_file_name: str,
    color1: tuple,
    color2: tuple,
    first_name: str,
    last_name: str,
    role: str,
    is_staff: bool,
    save_file: bool = True
):
    """
    It takes in a background image, two colors, a first name, a last name, and a
    subtitle, and returns an image with the first name and last name in the first
    color and the subtitle in the second color

    Args:
      background_file_name: the name of the background file
      color1: The color of the first line of text
      color2: The color of the text
      first_name: First name
      last_name: Last name
      role: The text that will be displayed on the image
      is_staff: Indicates if the image is for staff format

    Returns:
      The background image with the text added to it.
    """

    image = Image.open(background_file_name)
    add_text(image, color1, first_name, last_name,
             TITLE_FONT, 45, is_staff, True)
    add_text(image, color2, role, '', TEXT_FONT, 25, is_staff, False)
    file_name = f'{OUTPUT_DIR}/Badge {first_name} {last_name}.png'
    if save_file:
        image.save(file_name)
    return image


def create_badges_by_type(type: str, save_file = True):
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
    images = []
    for r in range(1, nb_rows):
        first_name = sheet.cell_value(rowx=r, colx=0)
        last_name = sheet.cell_value(rowx=r, colx=1)
        role = sheet.cell_value(rowx=r, colx=2)
        image = create_image(
            background_file_name,
            primary_color,
            SECONDARY_COLOR,
            first_name,
            last_name,
            role,
            type == "staff",
            save_file
        )
        images.append(image)
    return images


def create_all_badges(save_file = True):
    """Create all the badges from the excel document"""

    images = []
    for type in TYPES_SETTINGS:
        images += create_badges_by_type(type, save_file)
    return images


def create_one_badge():
    """Create one badge with custom data."""

    first_name = input('Prénom: ')
    last_name = input('Nom en Majuscule: ')
    print("Choisissez le type de badge à créer :")
    for type in TYPES_SETTINGS:
        print(f"\t* {type}")
    # lire le choix
    type = input('Entrez le type choisi : ')
    if type not in TYPES_SETTINGS:
        raise Exception("CHOIX INVALIDE !")
    # définir les paramètres
    settings = TYPES_SETTINGS[type]
    primary_color = settings["color"]
    background_file_name = settings["bg"]
    role = ""
    if type == "staff":  # Equipe Organisatrice
        role = input('Entrez le rôle: ')
    if type == "entreprise":  # Entreprise
        role = input('Entrez le nom de l’entreprise : ')
    # créer l'image
    image = create_image(
        background_file_name,
        primary_color,
        SECONDARY_COLOR,
        first_name,
        last_name,
        role,
        type == "staff"
    )
    return image


# script uniquement appelé lorsque ce fichier est directement exécuté
if __name__ == '__main__':
    print("Choisissez le type de badges à créer à partir du fichier source :")
    for type in TYPES_SETTINGS:
        print(f"\t* {type}")
    print("\t* all")
    print("\t* custom")
    input_cmd = input("Entrez votre choix: ")
    if input_cmd == "custom":
        create_one_badge()
    elif input_cmd == "all":
        create_all_badges()
    elif input_cmd in TYPES_SETTINGS:
        create_badges_by_type(input_cmd)
    else:
        raise Exception("CHOIX INVALIDE !")
    print(f"Badges créés dans le dossier '{OUTPUT_DIR}'")
