
from selenium import webdriver
from openpyxl import load_workbook 
import time
from selenium.webdriver.common.by import By


driver = webdriver.Chrome(executable_path=r"C:\Program Files (x86)\chromedriver.exe")
driver.get("http://127.0.0.1:5501/formulario.html")
time.sleep(3)

#VARIABLE PARA LA RUTA DEL EXCEL

filesheet = "./ejemplo.xlsx"

#CARGAR EL WORKBOOK

wb = load_workbook(filesheet)

#TRAER NOMBRE DE LAS HOJAS QUE ESTAN DISPONIBLES
hojas = wb.get_sheet_names()
print(hojas)

#VARIABLE PARA TOMAR LA HOJA
nombres = wb.get_sheet_by_name('Hoja 1')
wb.close()

for i in range(1,5): #recorre las filas
	nomb, apell, edad = nombres[f'A{i}:C{i}'][0] #recorrer las columnas
	print(nomb.value, apell.value, edad.value)
	time.sleep(1)
	driver.find_element(By.ID,"nom").send_keys(nomb.value)
	time.sleep(1)
	driver.find_element(By.ID,"ape").send_keys(apell.value)
	time.sleep(1)
	driver.find_element(By.ID,"edad").send_keys(str(edad.value))
	time.sleep(1)
	driver.find_element(By.ID,"enviar").click()
	time.sleep(1)
	print('--- Datos enviados ---')

driver.close() 