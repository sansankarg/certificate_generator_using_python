import pandas as pd
from PIL import Image, ImageDraw, ImageFont

#getting data from excel sheet
data = pd.read_excel('Book1.xlsx')

#setting positions and fonts for the contents edited in certificate
nameloc = (639, 520)
idloc = (807, 590)
dateloc = (195, 930)
session1loc = (822, 712)
session2loc = (954, 712)
subjectloc = (742, 670)
text_color = (1, 1, 1)
name_color = (225, 30, 38)
namefont = ImageFont.truetype("arialbd.ttf", 40)
idfont = ImageFont.truetype("arial.ttf", 25)
datefont = ImageFont.truetype("arial.ttf", 20)
sessionfont = ImageFont.truetype("arial.ttf", 18)
subjectfont = ImageFont.truetype("arialbd.ttf", 20)
python = Image.open("python.png")
datascience = Image.open("datascience.png")


for i, row in data.iterrows():
    #opening base certificate iamge file
    im = Image.open("certificate4.jpg")

    #creating draw function
    d = ImageDraw.Draw(im)

    #getting the data of individual
    name = row['can_name']
    canid = str(row['can_id'])
    session1 = row['date'].strftime("%B - %Y")
    session2 = row['cansession'].strftime("%B - %Y")
    date = row['date'].strftime("%d / %m / %Y")
    subject = row['subject']

    #selecting course logo according to individual course
    if subject == "PYTHON":
        logo = python.resize((150, 150))
    elif subject == "DATA SCIENCE":
        logo = datascience.resize((250,150))

    #editing the certificate
    im.paste(logo, (1272, 106))
    d.text(nameloc, name, fill = name_color, font = namefont)
    d.text(idloc, canid, fill = text_color, font = idfont)
    d.text(dateloc, date, fill = text_color, font = datefont)
    d.text(session1loc, session1, fill = text_color, font = sessionfont)
    d.text(session2loc, session2, fill = text_color, font = sessionfont)
    d.text(subjectloc, subject, fill = name_color, font = subjectfont)

    #saving and showing the final copy
    im.show()
    im.save("certificate_" + name + ".pdf")
