from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from create_badges import create_all_badges
import os

OUTPUT_FILE = "badges_FA.pdf"

# Be sure to set the correct current directory
os.chdir(os.path.dirname(__file__))

# Reportlab uses the "point" unit. We convert it to cm for usability :
PAGE_WIDTH, PAGE_HEIGHT = A4
CM = PAGE_WIDTH / 21  # ratio between points and centimeters

# generates all the images
badges = create_all_badges(save_file = False)

# badges images are 10cm x 6.5cm, including 0.5cm on each side for lost background
IMG_WIDTH = 10*CM
IMG_HEIGHT = 6.5*CM

# draw the PDF - to know more about how to add texts and other stuff :
# https://docs.reportlab.com/reportlab/userguide/ch2_graphics/#the-tools-the-draw-operations

pdf = canvas.Canvas(OUTPUT_FILE, A4)
for i in range(len(badges)):
    badge = badges[i]
    pos_x = 0.5*CM + (i % 2) * 10*CM
    pos_y = PAGE_HEIGHT - 0.7*CM - ((i % 8) // 2 + 1) * 7*CM
    pdf.drawInlineImage(badge, pos_x, pos_y, IMG_WIDTH, IMG_HEIGHT, True)
    if i % 8 == 7:
        pdf.showPage()

pdf.save()
