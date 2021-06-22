from posixpath import normcase
import pytesseract
from pytesseract import Output
import shutil
import random
try:
 from PIL import Image
except ImportError:
 import Image
 
import glob
import numpy as np
import cv2
import os
import shutil
import xlsxwriter
import re



pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'


def get_grayscale(image):
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

def thresholding(image):
    return cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

def leerfactura(image):
    img_org = image
    y=10
    x=10
    w=10
    h=10
    gray = get_grayscale(image)
    thresh = thresholding(gray)
    d_thresh = pytesseract.image_to_data(thresh, output_type=Output.DICT)
    keys = list(d_thresh.keys())
    mask = np.zeros((1755,1240), dtype="uint8")
    date_pattern = '[A-Z]+[A-Z]+[A-Z]+[A-Z]+[0-9]+[0-9]+[0-9]'
    n_boxes = len(d_thresh['text'])
    for i in range(n_boxes):
      if int(float(d_thresh['conf'][i])) > 60:
          if re.match(date_pattern, d_thresh['text'][i]):
            (x, y, w, h) = (d_thresh['left'][i], d_thresh['top'][i], d_thresh['width'][i], d_thresh['height'][i])
            img_data = cv2.rectangle(thresh, (x, y), (x + w, y + h), (0, 255, 0), 2)
  
    mask[y-5:y+h+5,x-5:x+w+5] = 255
    imgmasked=cv2.bitwise_and(img_org,img_org,mask = mask)
    imageOut = img_org[y-5:y+h+5,x-5:x+w+5]
    NroFactura = pytesseract.image_to_string(imageOut)
  
  
    keys = list(d_thresh.keys())
    mask = np.zeros((1755,1240), dtype="uint8")
    date_pattern = '(EMISION)'
    n_boxes = len(d_thresh['text'])
    for i in range(n_boxes):
      if int(float(d_thresh['conf'][i])) > 60:
          if re.match(date_pattern, d_thresh['text'][i]):
              (x, y, w, h) = (d_thresh['left'][i], d_thresh['top'][i], d_thresh['width'][i], d_thresh['height'][i])
              img_data = cv2.rectangle(thresh, (x, y), (x + w, y + h), (0, 255, 0), 2)
    mask[y-5:y+h+5,x+w:x+w+100] = 255
    img_org = image
    imgmasked=cv2.bitwise_and(img_org,img_org,mask = mask)
    imageOut = img_org[y-5:y+h+5,x+w:x+w+100]
    Fecha = pytesseract.image_to_string(imageOut)
  
    img_org = image
    imageOut = img_org[790:815,1040:1140]
    Total = pytesseract.image_to_string(imageOut)
    numero = Total.split(',')
    numero2 = numero[2].split('\n')
    numero3 = numero[0] + numero[1] + numero2[0]
    numero4 = numero3.split('.')
    Total=float(numero4[0])
  
    return NroFactura, Fecha, Total

def crearHojaCalculo(path):
    facturas = []
    expenses = ()
    print("Se imprime factura")
    print(path[0])
    for myPath in path:
        print("Otras facturas")
        print(myPath)
        nombre = myPath.split('/')
        nombre = nombre[len(nombre)-1]
        print("Movi"+nombre)
        finalNombre=".\Facturas\'"+nombre
        print(finalNombre)
        shutil.move(myPath,"./Facturas/"+nombre)

    files = glob.glob (".\Facturas\*.jpg")

    for myFile in files:
        print("Files")
        print(myFile)
        image = cv2.imread (myFile)
        facturas.append (image)
    i=2
    for factura in facturas:
        prueba_rgb = cv2.cvtColor(factura, cv2.COLOR_BGR2RGB)
        print("Antes Facutra")
        Factura, fecha, total=leerfactura(prueba_rgb)
        print("Factura")
        print(Factura)
        print(fecha)
        print(total)
        expenses += (
        [Factura, fecha, '900940013-3', 'Logística Roldan Garzón S.A.S', '=G'+str(i)+'/1.19', '=E'+str(i)+'*0.19', total],
        )
        i+=1

    workbook = xlsxwriter.Workbook('ComprasFacturas.xlsx')
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True})
    money = workbook.add_format({'num_format': '$#,##0'})

    worksheet.write('A1', '# Factura', bold)
    worksheet.write('B1', 'Fecha', bold)
    worksheet.write('C1', 'Nit', bold)
    worksheet.write('D1', 'Razón Social', bold)
    worksheet.write('E1', 'Valor Antes del Iva', bold)
    worksheet.write('F1', 'Iva', bold)
    worksheet.write('G1', 'Total', bold)

    row = 1
    col = 0

    for factura, fecha, nit, razonsocial, valorantes, iva, total in (expenses):
        worksheet.write(row, col,     factura)
        worksheet.write(row, col + 1, fecha)
        worksheet.write(row, col + 2, nit)
        worksheet.write(row, col + 3, razonsocial)
        worksheet.write(row, col + 4, valorantes)
        worksheet.write(row, col + 5, iva)
        worksheet.write(row, col + 6, total)
        row += 1

    workbook.close()
