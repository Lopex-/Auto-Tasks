import time
import os, sys





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

def contiene_texto(cadena):
	if cadena is None:
		return False
	elif cadena == "":
		return False
	else:
		return True

