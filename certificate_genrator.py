import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib

data = pd.read_excel (r'example.xlsx') #path of excel file
name_list = data["Full Name"].tolist()
for name in (name_list):
    im = Image.open(r'image.png')#add the image file name / certificate template 
    d = ImageDraw.Draw(im)
    location = (844, 510) # center point of where the name will be there
    # check example image and put the name in the correct place
    # you can use Microsoft Paint to find the correct location
    text_color = (69, 69, 69) #font color in RGB
    font = ImageFont.truetype("arial.ttf", 55) #font style and size
    text_width, text_height = font.getsize(name)
    text_x = location[0] - text_width / 2
    text_y = location[1] - text_height / 2
    text_location = (text_x, text_y)
    d.text(text_location, name, fill=text_color, font=font)
    im.save(f'example_{name}.pdf')#name of the pdf file
    print(f"Certificate for {name} has been created")