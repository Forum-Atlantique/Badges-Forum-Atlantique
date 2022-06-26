#pip install reportlab
from reportlab.pdfgen import canvas

from PIL import Image, ImageDraw, ImageFont

fileName = 'BadgesFA.pdf'
img = 'C:/Users/User/Downloads/test.jpg'
image = Image.open(r'C:/Users/User/Downloads/ScriptFABadges/backgroundChauffeur.png')

#pip install PyPDF2
from PyPDF2 import PdfFileReader
input1 = PdfFileReader(open('BadgesFA.pdf', 'rb'))
#à taper dans le terminal
[a,b,X,Y] = input1.getPage(0).mediaBox
#réponse
#RectangleObject([0, 0, 595.2756, 841.8898])
#[0, 0, width, height]
#WIdth = largeur


#Creating pdf object
pdf = canvas.Canvas(fileName)

#567
#400
w,h = image.size
image = image.resize((w//2,h//2))
#Attention ça diminue la résolution
#image.save('test.jpg')
#image.save('test.jpg', dpi =(200,200))
x = 0#coin de l'image sur le bord gauche
y = round(Y)- h//2 #coin de l'image en haut à gauche
hx=round(X) - w//2
hy = round(Y)-h
#pdf.drawImage(img, x, y, width=None, preserveAspectRatio=True, mask='auto')
#pdf.drawImage(img, x+h, y+h, width=None, preserveAspectRatio=True, mask='auto')
pdf.drawInlineImage(image, x, y, width = None, height= None, preserveAspectRatio = True)
#pdf.drawInlineImage(image, x+h, y, width = None, height= None, preserveAspectRatio = True)
pdf.save()

#canvas.drawInlineImage(self, image, x,y, width=None,height=None)
#canvas.drawImage(self, image, x,y, width=None,height=None,mask=None)
#The showPage method finishes the current page. All additional drawing will be done on another page.

