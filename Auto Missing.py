#=========================================================================================#   
#                                                                                         #
#                          Grandioso programa para cargar missings                        #
#                                   Ver 0.3 - 22/09/2018                                  #
#                                    By. Irvin R. Lópex                                   #
#                                                                                         #
#=========================================================================================#
#                                   Ver 0.0 - 22/09/2018                                  #
# En realidad, el codigo acaba de pasar a una especie de Beta. No se ha probado           #
# para cargar missings en si, solo es una base que considero suficiente para desarrollar  #
# el programa deceado. Puede cambiar entre pedidos en SF, abrir los cuadros de texto      #
# de Production Notes y Material Status, y guardar los datos. Es lo único que se ha       #
# probado. Se concluyen las pruebas pre-natales, este proyecto acaba de nacer.            #
#                                                                                         #
#=========================================================================================#
#                                   Ver 0.1 - 16/10/2018                                  #
# El proyecto ya ha cargado missings un par de veces. Hay muchisimas cosas que reparar,   #
# entre ellas la "v" que se pone antes de pegar los valores en los campos de texto.       #
#                                                                                         #
#											CAMBIOS                                       #
# 1.- Se añadio un limitador de caracteres para evitar exceder el límite del campo de txt #
#                                                                                         #
#=========================================================================================#
#                                   Ver 0.2 - 08/11/2018                                  #
# Se cargan missings exitosamente. Aún hay ligeros detalles por reparar como cuando el    # 
# string tiene más del límite de caracteres y se corta a mitad de línea. No recuero que   #
# otros cambios hice.                                                                     #
#                                                                                         #
#=========================================================================================#
#                                   Ver 0.3 - 30/03/2019                                  #
# Por fin me deshice de la libreria clipboard. Ya se puede usar copiar y pegar mientras   #
# corre este programa. El código cambio a forma modular y se utilizan mas utilidades de   #
# python. Se añadió la función de elegir el archivo missing.xls* desde el explorador por  #
# si no se encuentra en la ruta proporcionada. Quedan pendientes la interfaz grafia,      #
# saber cuando la pki está o no conectada, cambiar los comentarios al formato de Python,  #
# poder tomar valores desde las formulas en excel, entre otras.                           #
#                                                                                         #
#=========================================================================================#
#                                  Ver 0.31 - 30/04/2019                                  #

#	Se añadieron muchas excepciones donde habia problemas poco recurrentes y algo escondidos
# se repararon bugs, se añadieron las fechas de Last material y PO Release, los archivos 
#de ayuda y chromedriver.exe se encuentran en la misma carpeta que este script,
# hubo una ligera modificación en la función cargar_xlsx donde recibe la dirección completa
# del archivo.
#=========================================================================================#


# Formato de columnas:
# | Factory SO | Sold To Party Name | Item Num | Faltantes | Encabezado | Texto Missing | Link | Last Material | PO Release |






#==========================================================================================
# 
#                                     Importar Cosas
#
#==========================================================================================
# Codigo tomado prestado de otras librerias. Gracias a sus desarrolladores, a las
# personas que hacen preguntas en foros, a las que las responden y a los que hacen
# tutoriales.
#==========================================================================================

#Usamos Selenium para comunicarnos con el headless browser, en este caso chrome
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, WebDriverException

#Usamos Openpyxl para la comunicación con archivos xlsx
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

#Usamos Time para obtener la hora de inicio y fin del programa
import time

#Con OS obtenemos la ruta del script 
import os, sys

#Para obtener la ruta del archivotl
import tkinter
from tkinter.filedialog import askopenfilename

import datetime




#==========================================================================================
# 
#                                     Ignorar warnings
#
#==========================================================================================
# Desconosco como funciona este código pero evita que aparescan advertencias durante
# la ejecucion.
#==========================================================================================

def warn(*args, **kwargs):
    pass
import warnings
warnings.warn = warn




#==========================================================================================
# 
#                                 Tomar tempo
#
#==========================================================================================
# Funcion que devuelve el tiempo (hora) actual 
#==========================================================================================

def tomar_tiempo():
	return time.time()




#==========================================================================================
# 
#                           Imprimir tiempo de ejecucion
#
#==========================================================================================
# Funcion que devuelve el tiempo (hora) actual 
#==========================================================================================

def imprimir_tiempo_ejecucion(tiempo_inicial, tiempo_final):
	total_segundos = round(tiempo_final - tiempo_inicial)
	minutos = round( (total_segundos/60) - 0.5 )
	segundos = total_segundos - ( minutos * 60 )
	print("Tiempo de ejecucion: ", minutos, " minutos con ", segundos, " segundos")
	os.system('pause')




#==========================================================================================
# 
#                                 Obtener esta ruta
#
#==========================================================================================
# Obtiene la ruta (Path) en la que se encuentra guardado el script.
#==========================================================================================

def obtener_esta_ruta():
	esta_ruta = os.path.dirname(sys.argv[0])
	return esta_ruta




#==========================================================================================
# 
#                                     Cargar excel
#
#==========================================================================================
# Abre y ancla el archivo de excel a una variable.
#==========================================================================================

def cargar_excel(nombre_archivo):
	try:
		archivo_xlsx = openpyxl.load_workbook(nombre_archivo)
		print("Archivo ", nombre_archivo, " cargado.")

	except FileNotFoundError:
		print("No se encontró el archivo ", nombre_archivo)
		tkinter.Tk().withdraw()
		nombre_archivo = askopenfilename()
		archivo_xlsx = cargar_excel(nombre_archivo)
	
	except InvalidFileException:
		print("Archivo no valido, intenta de nuevo")
		tkinter.Tk().withdraw()
		nombre_archivo = askopenfilename()
		archivo_xlsx = cargar_excel(nombre_archivo)

	return archivo_xlsx




#==========================================================================================
# 
#                                     Cargar hoja
#
#==========================================================================================
# Abre y ancla una hoja del archivo excel a una variable.
#==========================================================================================

def cargar_hoja(archivo_excel, nombre_hoja):
	try:
		hoja_excel = archivo_excel.get_sheet_by_name(nombre_hoja)
		print("Hoja ", nombre_hoja, "cargada.")

	except KeyError:
		print("No se encontró la hoja ", nombre_hoja, " en el archivo excel")
		nombre_hoja = input("Ingersar el nombre de la hoja de excel: ")
		print(nombre_hoja)
		hoja_excel = cargar_hoja(archivo_excel, nombre_hoja)

	return hoja_excel




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def crear_explorador():
	chromeOptions = webdriver.ChromeOptions()
	chromeOptions.add_experimental_option('useAutomationExtension', False)
	ruta_chromedriver = obtener_esta_ruta() + "\\chromedriver.exe"
	driver = webdriver.Chrome(ruta_chromedriver, chrome_options=chromeOptions, desired_capabilities=chromeOptions.to_capabilities())
	return driver




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def objetos_de_salesforce(explorador, intento=0):
	try:
		objetos = { "txt_field_material_status" : explorador.find_element_by_id("00NA000000Ao5rT_ileinner"),
				"txt_field_production_notes" : explorador.find_element_by_id("00NA000000AAcgL_ilecell"),
				"boton_save" : explorador.find_element_by_name("inlineEditSave"),
				"boton_cancel" : explorador.find_element_by_name("inlineEditCancel"),
				 "txt_field_last_material" : explorador.find_element_by_id("00NA000000AoxwQ_ileinner"),
				 "txt_field_PO_release" : explorador.find_element_by_id("00NA000000AAcfL_ileinner"),}
	except NoSuchElementException:
		if intento == 0:
			time.sleep(5)
			objetos = objetos_de_salesforce(explorador, 1)
		else:
			print("No se encuentran los objetos en la pagina. Cargar manualmente la pagina del proyecto actual.")
			os.system('pause')
			objetos = objetos_de_salesforce(explorador, 2)

	return objetos




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def cargar_link(explorador, link):
	explorador.get(link)
	time.sleep(1)




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def ayuda():

	try:
		ruta_ayuda = obtener_esta_ruta() + "\\chromedriver.dll" 
		archivo_texto = open(ruta_ayuda, 'r')
		texto_ayuda = archivo_texto.read()
		print (texto_ayuda)
		archivo_texto.close()

	except FileNotFoundError:
		print("No se encontró el archivo de ayuda.")

	

#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def elegir_modo():
	print("")
	modo=input("Ingrese 1 para ejecutar y guardar cambios al finalizar, o cualquier otro número para hacer una prueba: ")
	if modo.lower() == "ayuda" or modo.lower() == "help" :
		ayuda()
		modo = elegir_modo()
	return modo




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def iniciar_sesion_siemens(explorador):
	cargar_link(explorador, "https://siemens.my.salesforce.com/home/home.jsp")
	print("Inicia sesion con tu PKI...")
	os.system('pause')




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def columnas_de_missing():
	columnas = { "Factory_SO" : 1,
				#"Sold_to_party_name" : 2,
				#"Item_Num" : 3,
				"Link" : 7,
				"Faltantes" : 2,
				"Texto_missing" : 3,
				"Encabezado" : 4,
				"Last_material" : 5,
				"PO_release" : 6,}
	return columnas




#==========================================================================================
# 
#                                 
#
#==========================================================================================
# 
#==========================================================================================

def actualizar_celdas(hoja, fila, columnas):
	celda = {}
	for nombre, columna in columnas.items():
		celda.update( {nombre : hoja.cell(row = fila, column = columna)} )
	return celda




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def es_un_link(link):
	if link is None:
		return False
	elif "https://siemens.my.salesforce.com" not in link:
		return False
	else:
		return True 




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def recortar_a_n_caracteres(texto, max_caracteres):
	if len(texto) > max_caracteres:
		texto = texto[0:max_caracteres]
	return texto




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def cargar_material_status(explorador, txt_field_material_status, texto_missing):
	texto_missing = recortar_a_n_caracteres(texto_missing, 1800)
	
	try:
		txt_field_material_status.click()
	except WebDriverException:
		print("Elemento desenfocado. No utilizar el explorador.")
		os.system('pause')
		txt_field_material_status.click()

	webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()
	time.sleep(1)
	webdriver.ActionChains(explorador).send_keys(texto_missing).perform()
	time.sleep(1)
	webdriver.ActionChains(explorador).send_keys(Keys.TAB).perform()
	webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def contiene_texto(cadena):
	if cadena is None:
		return False
	elif cadena == "":
		return False
	else:
		return True




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def cargar_production_notes(explorador, txt_field_production_notes, encabezado):
	if contiene_texto(encabezado):
		try:
			txt_field_production_notes.click()
		except WebDriverException:
			print("Elemento desenfocado. No utilizar el explorador.")
			os.system('pause')
			txt_field_production_notes.click()

		webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()
		time.sleep(1)
		webdriver.ActionChains(explorador).key_down(Keys.CONTROL).perform()
		webdriver.ActionChains(explorador).send_keys(Keys.HOME).perform()
		webdriver.ActionChains(explorador).key_up(Keys.CONTROL).perform()
		webdriver.ActionChains(explorador).send_keys(encabezado).perform()
		webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()
		time.sleep(1)
		webdriver.ActionChains(explorador).send_keys(Keys.TAB).perform()
		webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()






def convertir_fecha_a_str(fecha):
	if type(fecha) == datetime.datetime:
		nueva_fecha = fecha.strftime('%d/%m/%Y')
	else:
		nueva_fecha = fecha	
	return nueva_fecha






#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def cargar_last_material(explorador, txt_field_last_material, fecha_last_material):
	
	str_fecha_last_material = convertir_fecha_a_str(fecha_last_material)

	if contiene_texto(str_fecha_last_material):
		try:
			txt_field_last_material.click()
		except WebDriverException:
			print("Elemento desenfocado. No utilizar el explorador.")
			os.system('pause')
			txt_field_last_material.click()

		webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()
		time.sleep(1)
		webdriver.ActionChains(explorador).send_keys(str_fecha_last_material).perform()
		time.sleep(1)
		webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def cargar_PO_release(explorador, txt_field_PO_release, fecha_PO_release):
	
	str_fecha_PO_release = convertir_fecha_a_str(fecha_PO_release)
	
	if contiene_texto(str_fecha_PO_release):
		try:
			txt_field_PO_release.click()
		except WebDriverException:
			print("Elemento desenfocado. No utilizar el explorador.")
			os.system('pause')
			txt_field_PO_release.click()

		webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()
		time.sleep(1)
		webdriver.ActionChains(explorador).send_keys(str_fecha_PO_release).perform()
		time.sleep(1)
		webdriver.ActionChains(explorador).send_keys(Keys.ENTER).perform()




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def imprimir_estado_missing(celda):
	if celda["Faltantes"].value == 1:
		print(celda["Factory_SO"].value," - ", celda["Faltantes"].value, " faltante" )
	else:
		print(celda["Factory_SO"].value," - ", celda["Faltantes"].value, " faltantes" )




#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def cargar_missing(explorador, hoja_missing, modo):
	print("Iniciando proceso de actualizar datos en salesforce.")
	fila_inicial = 2
	fila_actual = fila_inicial
	columna = columnas_de_missing()
	celda = actualizar_celdas(hoja_missing, fila_actual, columna)
	acciones = ActionChains(explorador)

	while es_un_link(celda["Link"].value) :
		imprimir_estado_missing(celda)

		cargar_link(explorador, celda["Link"].value)
		objeto = objetos_de_salesforce(explorador)

		cargar_production_notes(explorador, objeto["txt_field_production_notes"], celda["Encabezado"].value)
		cargar_material_status(explorador, objeto["txt_field_material_status"], celda["Texto_missing"].value)
		cargar_last_material(explorador,objeto["txt_field_last_material"], celda["Last_material"].value)
		cargar_PO_release(explorador,objeto["txt_field_PO_release"], celda["PO_release"].value)

		if modo == "1":
			objeto["boton_save"].click()
		else:
			objeto["boton_cancel"].click()

		fila_actual += 1
		celda = actualizar_celdas(hoja_missing, fila_actual, columna)
		time.sleep(4)
	print(fila_actual - fila_inicial, " proyectos actualizados")




#==========================================================================================
# 
#                                 Programa principal
#
#==========================================================================================
# Eh aqui el inicio del codigo a ejecutarts.
#==========================================================================================
def main():
	tiempo_inicial = tomar_tiempo()

	modo = elegir_modo()

	xlsx_missing = cargar_excel( obtener_esta_ruta() + "\\Auto-missing.xlsx")

	hoja_missing = cargar_hoja(xlsx_missing, "SF")
	
	explorador = crear_explorador()

	iniciar_sesion_siemens(explorador)

	cargar_missing(explorador, hoja_missing, modo)

	explorador.quit()

	tiempo_final = tomar_tiempo()

	imprimir_tiempo_ejecucion(tiempo_inicial, tiempo_final)




if __name__ == "__main__":
	main()