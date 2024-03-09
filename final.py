from PyPDF2 import PdfWriter, PdfReader

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
import reportlab.rl_config
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import io
import pandas as pd

pdfmetrics.registerFont(TTFont('Noyh Slim R', 'font/NoyhSlimR-Regular.ttf'))
pdfmetrics.registerFont(TTFont('VAG Rounded BT', 'font/vag_rounded_bt.ttf'))

participants = pd.read_excel('lista.xlsx')
participants

# Funcion que dependiendo de la longitud de la palabra, el tamaño de la font varia.
def calcular_tamaño_fuente(palabra):
    longitud_palabra = len(palabra)
    if longitud_palabra <= 5:
        return 50
    elif longitud_palabra <= 10:
        return 40
    else:
        return 35

for i in range(len(participants)):
    
    ### Crea variables que van a ser añadidas al PDF
    codigo = participants.loc[i,'Cod']
    producto = participants.loc[i,'Producto']
    precio = participants.loc[i,'Precio']
    Descuento = participants.loc[i,'Descuento']
    Dia = participants.loc[i, 'Dia']
    SegundaU = participants.loc[i,'SegundaU']
    
    ### Crear el tamaño de la pagina a editar
    packet = io.BytesIO()
    width, height = letter 
    c = canvas.Canvas(packet, pagesize=(width*2, height*2)) 
    
    ### Se establecen las fonts que van a llevar las variables en el PDF
    reportlab.rl_config.warnOnMissingFontGlyphs = 0
    pdfmetrics.registerFont(TTFont('Noyh Slim R', 'font/NoyhSlimR-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('VAG Rounded BT', 'font/vag_rounded_bt.ttf'))
    
    ### Se establece el color, el tipo letra y tamaño y ademas en donde va a estar situada. 
    # variable 1 : codigo
    c.setFillColorRGB(0,0,0) #Se coloca el color que va a llevar 
    c.setFont('Noyh Slim R', 30)  #Se coloca el tipo le letra y el tamaño  
    c.drawCentredString(535,800, str(codigo))    #Coordenadas donde va a ir ubicado el texto
    
    # variable 2 : producto
    c.setFillColorRGB(0,0,0) 
    c.setFont('VAG Rounded BT', calcular_tamaño_fuente(producto))                     
    c.drawCentredString(300, 570, producto)     
    
    # variable 3 : precio 
    c.setFillColorRGB(0,0,0)     
    c.setFont('Noyh Slim R', 50)                   
    c.drawCentredString(450, 170, str(precio))        

    # variable 4 : Descuento
    c.setFillColorRGB(0,0,0)      
    c.setFont('Noyh Slim R', 180)                   
    c.drawCentredString(310, 320, "{:.0%}".format(Descuento))        

    # misma variable 4 pero colocada en la parte superior izquierda en blanco
    c.setFillColorRGB(255,255,255)     
    c.setFont('Noyh Slim R', 40)                   
    c.drawCentredString(80, 760, "{:.0%}".format(Descuento))        


    # variable 5 : Dia
    c.setFillColorRGB(255,255,255)    
    c.setFont('VAG Rounded BT', 25)                   
    c.drawCentredString(75, 800, str(Dia))        

    # variable 6 : Segunda U
    c.setFillColorRGB(0,0,0)      
    c.setFont('Noyh Slim R', 50)                   
    c.drawCentredString(450, 110, str(SegundaU))        

    # Save all Canvas settings
    c.save()

    ### Step 6: Get final PDF certifciate
    # Get PDF template
    existing_pdf = PdfReader(open("cartel.pdf", "rb"))   # Lee el PDF
    page = existing_pdf.pages[0]              # Toma la primera pagina
    
    # Inserta los textos en el PDF  
    packet.seek(0)                            
    new_pdf = PdfReader(packet)               # Crea un nuevo pdf 
    page.merge_page(new_pdf.pages[0])         # Agrega los textos al nuevo pdf 

    ### Exporta el nuevo PDF
    file_name = producto.replace(" ","_")
    certificado = "Certificates/" + file_name + "_frioteka.pdf"
    outputStream = open(certificado, "wb")
    output = PdfWriter()
    output.add_page(page)
    output.write(outputStream)
    outputStream.close()