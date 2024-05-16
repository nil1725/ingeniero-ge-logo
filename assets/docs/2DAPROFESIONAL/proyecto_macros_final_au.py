# -*- coding: utf-8 -*-
"""
Editor de Spyder

Este es un archivo temporal
"""
import xlwings as xw
import pandas as pd
# from pywinauto import Application

# Ruta al archivo Excel de QA/QC
archivo_excel_qaqc = r"D:/Documentos\Escritorio\NILSON\Reporte_Resumen_QAQC\INMACULADA NUEVO\QC Inmaculada\qaqc_canales.xlsx"

# Leer el archivo Excel en un DataFrame de pandas
df = pd.read_excel(archivo_excel_qaqc)

# Mostrar las primeras filas del DataFrame para tener una idea de su estructura
#print(df.head())

# Filtrar el DataFrame para las filas donde el valor en la columna 'Interlab' no es 'Interlab'
df_interlab= df[df['Interlab'] != 'Interlab']

filtro_estandares = df_interlab[df_interlab['TipoMuestra'] == 'Estándar']

tipo_estandar = filtro_estandares[filtro_estandares['Descripcion'] == 'IN-16'].reset_index(drop=True)

# Extraer los valores de la columna "Despacho"
valores_despacho = tipo_estandar['Despacho']
# Extraer los valores de la columna "Despacho"
valores_certificado = tipo_estandar['Certificado']
valores_fecha=tipo_estandar["fecha_reporte"]
valores_muestra=tipo_estandar["vMuestraControl"]
valores_au=tipo_estandar["Au_ppm"]
#print(valores_au)

# Ruta al archivo Excel que contiene la macro
archivo_excel = r"D:\Documentos\Escritorio\NILSON\Reporte_Resumen_QAQC\INMACULADA NUEVO\QC Inmaculada\canales\CN Estándares IN-16 AU.xlsm"

# Iniciar una instancia de Excel y abrir el libro
app = xw.App()
wb = app.books.open(archivo_excel)

# Ejecutar la macro
wb.macro('CLEAR')()

# app_xl = Application().connect(title='Microsoft Excel')
# dlg = app_xl.top_window()
# dlg['Button'].click()
# Seleccionar la hoja "INPUT"
hoja_input = wb.sheets['INPUT']

# Pegar los valores en la hoja "INPUT" fila por fila
for i in range(len(valores_despacho)):
    hoja_input.range(f'A{i+11}').value = valores_despacho[i]
    hoja_input.range(f'B{i+11}').value = valores_certificado[i]
    hoja_input.range(f'C{i+11}').value = valores_fecha[i]
    hoja_input.range(f'D{i+11}').value = valores_muestra[i]
    hoja_input.range(f'E{i+11}').value = valores_au[i]

# Definir el rango de destino para cada columna
# rango_despacho = hoja_input.range(f'A11:A{10 + len(valores_despacho)}')
# rango_certificado = hoja_input.range(f'B11:B{10 + len(valores_certificado)}')
# rango_fecha = hoja_input.range(f'C11:C{10 + len(valores_fecha)}')
# rango_muestra = hoja_input.range(f'D11:D{10 + len(valores_muestra)}')
# rango_au = hoja_input.range(f'E11:E{10 + len(valores_au)}')

# # Pegar los valores en los rangos de destino
# rango_despacho.value = valores_despacho.values.tolist()
# rango_certificado.value = valores_certificado.values.tolist()
# rango_fecha.value = valores_fecha.values.tolist()
# rango_muestra.value = valores_muestra.values.tolist()
# rango_au.value = valores_au.values.tolist()
 
# Guardar cualquier cambio realizado en el libro
wb.save()

# Cerrar el libro de Excel
wb.close()

# Cerrar la instancia de Excel
app.quit()

