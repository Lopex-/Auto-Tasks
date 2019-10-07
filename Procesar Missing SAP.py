#ver 0.1
#Da formato bien chido.
#Ver 0.1.1 - 22-05-2019
#Elimino un salto de linea al final del encabezado.




from xlsx import cargar_excel, cargar_hoja, actualizar_celdas, guardar_xlsx
from fechas import hoy, convertir_str_a_fecha, convertir_fecha_a_str
from utilidad import tomar_tiempo, imprimir_tiempo_ejecucion, obtener_esta_ruta



#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================


def cargar_columnas_sap():
	columnas = { "Secuencia" : 1, 
				"Sales_Ord" : 2,
				"Dates_Repl" : 3,
				"Material" : 4,
				"Description" : 5,
				"Requiremen" : 7,
				"Unidad" : 9,
				"Proveedor" : 10,}
				#"Repl_Elem" : 11,
				#"Requ_Date" : 12,
				#"IssSlo" : 13,
				#"Proc" : 14,
				#"Supplier_Country" : 15,
				#"Creation_Date" : 16,
				#"MRP" : 17,}
	return columnas




def cargar_columnas_sf():
	columnas = { "Sales_Ord" : 1, 
				"Faltantes" : 2,
				"Material_Status" : 3,
				"Production_Notes" : 4,
				"Last_Material" : 5,
				"PO_Release" : 7,
				"Link" : 9,}
	return columnas



#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def celda_no_vacia(sales_order_actual):
	if sales_order_actual is None:
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

def obtener_texto_material(celdas_sap):
	texto = ""
	
	if celda_no_vacia(celdas_sap["Dates_Repl"].value):
		if type(celdas_sap["Dates_Repl"].value) == str:
			texto += celdas_sap["Dates_Repl"].value
		else:
			texto += celdas_sap["Dates_Repl"].value.strftime("%m/%d/%Y")
	texto += " "




	if celda_no_vacia(celdas_sap["Material"].value):
		texto += celdas_sap["Material"].value
	texto += " "


	if celda_no_vacia(celdas_sap["Description"].value):
		texto += celdas_sap["Description"].value
	texto += " "


	if celda_no_vacia(celdas_sap["Requiremen"].value):
		str_requi = str(celdas_sap["Requiremen"].value)
		texto += str_requi
	texto += " "


	if celda_no_vacia(celdas_sap["Unidad"].value):
		texto += celdas_sap["Unidad"].value
	texto += " "


	if celda_no_vacia(celdas_sap["Proveedor"].value):
		texto += celdas_sap["Proveedor"].value
	

	return texto















#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def agrupar_faltantes(hoja_sap):
	proyectos = {}
	columnas_sap = cargar_columnas_sap()
	fila_actual = 2
	celdas_sap = actualizar_celdas(hoja_sap, fila_actual, columnas_sap)

	while celda_no_vacia(celdas_sap["Sales_Ord"].value):
		
		if celdas_sap["Sales_Ord"].value not in proyectos:
			proyectos.update( { celdas_sap["Sales_Ord"].value : [] } )
		
		str_material = obtener_texto_material(celdas_sap)
		proyectos[celdas_sap["Sales_Ord"].value].append(str_material)

		fila_actual+=1
		celdas_sap=actualizar_celdas(hoja_sap, fila_actual, columnas_sap)
	return proyectos







#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def obtener_last_material(texto_missing):
	lineas = texto_missing.splitlines()
	peor_fecha = hoy()

	for linea in lineas:
		
		palabras = linea.split()
		str_fecha = palabras[0]
		
		try:
			fecha = convertir_str_a_fecha(str_fecha)
			if fecha > peor_fecha:
				peor_fecha = fecha
		
		except ValueError:
			next

	if peor_fecha == hoy():
		return None
	else:
		str_peor_fecha = convertir_fecha_a_str(peor_fecha)
		return str_peor_fecha







#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def escribir_faltantes(proyectos, hoja_sf):
	fecha_hoy = hoy()
	str_hoy = fecha_hoy.strftime("%m/%d")

	fila_actual = 2
	columnas_sf = cargar_columnas_sf()
	celdas_sf = actualizar_celdas(hoja_sf, fila_actual, columnas_sf)

	for proyecto, faltantes in proyectos.items():
		
		texto_missing = ""
		caracteres = 0
		for faltante in faltantes:
			caracteres += len(faltante)

			if caracteres<1800:
				texto_missing += (faltante + "\n")
			else: 
				continue

		last_material = obtener_last_material(texto_missing)

		n_faltantes = str(len(faltantes))
		encabezado = str_hoy + ": " + n_faltantes + " MP"
		texto_missing = encabezado + "\n" + texto_missing

		celdas_sf["Sales_Ord"].value = proyecto
		celdas_sf["Faltantes"].value = n_faltantes
		celdas_sf["Production_Notes"].value = encabezado
		celdas_sf["Material_Status"].value = texto_missing
		celdas_sf["Last_Material"].value = last_material

		fila_actual += 1
		celdas_sf = actualizar_celdas(hoja_sf, fila_actual, columnas_sf)















#==========================================================================================
# 
#                                
#
#==========================================================================================
# 
#==========================================================================================

def main():
	xlsx_automis = cargar_excel(obtener_esta_ruta() + "\\Auto-Missing.xlsx")

	hoja_sap = cargar_hoja(xlsx_automis, "SAP")

	hoja_sf = cargar_hoja(xlsx_automis, "SF")

	proyectos = agrupar_faltantes(hoja_sap)

	escribir_faltantes(proyectos, hoja_sf)

	guardar_xlsx(xlsx_automis, obtener_esta_ruta() + "\\Auto-Missing")
	xlsx_automis.close()

	print("fin del programa")



if __name__ == "__main__":
	main()