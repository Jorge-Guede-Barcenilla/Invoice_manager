from io import StringIO
try:
    from string import maketrans   # Python 2
except ImportError:
    maketrans = str.maketrans      # Python 3
import re
import pyodbc
import os
import sys
from os import listdir
from os.path import isfile, join
import tkinter
from tkinter import Tk
from tkinter.filedialog import askdirectory
from tkinter import messagebox
import ctypes
import subprocess
import math

# Function to regex search
def ParseFields(t, u = '', v = '', w = '', x = '', y = '', z = '', field_name = "", occurrence = 1):

    # initialize regex search
    regex_search = t + u + v + w + x + y + z

    try:
        Invoice_search = re.search(regex_search, text)
        #print(Invoice_search)

        if Invoice_search:
            print(field_name + " ----> " + Invoice_search.group(occurrence))
            return(Invoice_search.group(occurrence))
        else:
            print(field_name + " -----> NULL")

    except AttributeError:
        return ("Error")

# Function to convert a list to string
def List2String(list):

    # initialize an empty string
    string = " "

    # return string
    return (string.join(list))

# Round up with five
def round_down(n, decimals=2):

    multiplier = 10 ** decimals
    print(str(math.floor(float(n.replace('.','DOT').replace('DOT','').replace(',','.')) * multiplier) / multiplier).replace('.', ','))
    return str(math.floor(float(n.replace('.','DOT').replace('DOT','').replace(',','.')) * multiplier) / multiplier).replace('.', ',')

# Check deviation.
def check_deviation_billed_power(lectura_real_maximetro, potencia_contratada, potencia_facturada_peaje_acceso, potencia_facturada_margen_comercializadora):
    print(lectura_real_maximetro)
    if lectura_real_maximetro != 0 and (lectura_real_maximetro < 0.85 * potencia_contratada):
        if (abs(0.85 * potencia_contratada - potencia_facturada_peaje_acceso) > 0.01) or (abs(0.85 * potencia_contratada != potencia_facturada_peaje_acceso) > 0.01):
           print("power billed incorrectly")
    elif lectura_real_maximetro != 0 and (lectura_real_maximetro > 1.05 * potencia_contratada):
        if (abs((potencia_contratada + 2 * (lectura_real_maximetro - 1.05 * potencia_contratada)) - potencia_facturada_peaje_acceso) > 0.01) or (abs((potencia_contratada + 2 * (lectura_real_maximetro - 1.05 * potencia_contratada)) - potencia_facturada_margen_comercializadora) > 0.01):
           print("power billed incorrectly")
    else:
        if ((lectura_real_maximetro - potencia_facturada_peaje_acceso > 0.01) or (lectura_real_maximetro - potencia_facturada_peaje_acceso > 0.01)):
           print("power billed incorrectly")

def check_deviation_product(total, w = "1", x = "1", dias_periodo = 1, dias_año = 1):

    if abs(float(total.replace('.','DOT').replace('DOT','').replace(',','.')) - (float(w.replace('.','DOT').replace('DOT','').replace(',','.')) * float(x.replace('.','DOT').replace('DOT','').replace(',','.')) * float(str(dias_periodo).replace('.','DOT').replace('DOT','').replace(',','.')) * 1/float(str(dias_año).replace('.','DOT').replace('DOT','').replace(',','.')))) > 0.02:
         print(float(w.replace('.','DOT').replace('DOT','').replace(',','.')) * float(x.replace('.','DOT').replace('DOT','').replace(',','.')) * float(str(dias_periodo).replace('.','DOT').replace('DOT','').replace(',','.')) * 1/float(str(dias_año).replace('.','DOT').replace('DOT','').replace(',','.')))
         print("Total cost: " + total + " is not consistent -----> " + w + " * " + x + " * " + str(dias_periodo) + " * 1/" + str(dias_año))


def check_deviation_sum(total, t = "0", u = "0", v = "0", w = "0", x = "0", y = "0", z = "0"):

    if abs(float(total.replace('.','DOT').replace('DOT','').replace(',','.')) - (float(t.replace('.','DOT').replace('DOT','').replace(',','.')) + float(u.replace('.','DOT').replace('DOT','').replace(',','.')) + float(v.replace('.','DOT').replace('DOT','').replace(',','.')) + float(w.replace('.','DOT').replace('DOT','').replace(',','.')) + float(x.replace('.','DOT').replace('DOT','').replace(',','.')) + float(y.replace('.','DOT').replace('DOT','').replace(',','.')) + float(z.replace('.','DOT').replace('DOT','').replace(',','.')))) > 0.02:
         print("Total cost: " + total + " is not consistent -----> " + t + " + " + u + " + " + v + " + " + w + " + " + x + " + " + y + " + " + z)

#-------------------SELECT DIRECTORY AND CONVERSION TO .TXT
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
okcancel = messagebox.askyesno('Conversion to txt', 'Do you want to perform this action?') # OK / Cancel

if okcancel is True:
    input_directory = askdirectory() # show an "Open" dialog box and return the path to the selected file

    if input_directory == '':
        ctypes.windll.user32.MessageBoxW(0, "Conversion to txt cancelled", "Conversion to txt", 1)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Conversion to txt accepted. Input directory: " + input_directory, "Conversion to txt", 1)
        files2convert = [f for f in listdir(input_directory) if isfile(join(input_directory, f))]
        files2convert_string = List2String(files2convert)

    output_directory = askdirectory() # show an "Open" dialog box and return the path to the selected file

    if output_directory == '':
        ctypes.windll.user32.MessageBoxW(0, "Conversion to txt cancelled", "Conversion to txt", 1)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Conversion to txt a accepted. Output directory: " + output_directory, "Conversion to txt", 1)

    if files2convert_string == '':
        ctypes.windll.user32.MessageBoxW(0, "Conversion to txt cancelled. No files found in directory.", "Conversion to txt", 1)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Conversion to txt accepted. List of paths: " + files2convert_string, "Conversion to txt", 1)
        # Execute ExtractText program through cmd to convert .pdf to .txt
        process = subprocess.Popen("C:/Users/Jorge/Desktop/PycharmProjects/Invoice_manager/ExtractText.exe", shell=True, stdout=subprocess.PIPE)
        process.wait()
        returncode = process.returncode

        for file in files2convert:
            command = "EXTRACTTEXT input=\"" + os.path.join(input_directory, file) + "\" output=\""
            command = command + os.path.join(output_directory, os.path.splitext(file)[0]) + ".txt\""
            process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
            process.wait()
            returncode = process.returncode

else:
     ctypes.windll.user32.MessageBoxW(0, "Conversion to txt cancelled", "Conversion to txt", 1)

#------------------- DEFINE INPUT DIRECTORY OF INVOICES
input_directory_invoices = askdirectory()
if input_directory_invoices == '':
    ctypes.windll.user32.MessageBoxW(0, "Execution cancelled", "Invoice processing", 1)
    sys.exit()
else:
    ctypes.windll.user32.MessageBoxW(0, "Invoice conversion process started. Input directory: " + input_directory_invoices, "Invoice processing", 1)
    files2process = [f for f in listdir(input_directory_invoices) if isfile(join(input_directory_invoices, f))]
    files2process_string = List2String(files2process)
    ctypes.windll.user32.MessageBoxW(0, "Processing of invoices accepted. List of invoices: " + files2process_string, "Invoice processing", 1)

#------------------- READ TXT FILES ITERATIVELY, PARSE ALL FIELDS, GENERATE REPORTS AND STORE IN DATABASE

# DATABASE CONNECTION
import mysql.connector

mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  passwd="MySQLFE16137007966831",
  database="Renowattio"
)

# READ TXT FILES ITERATIVELY
for i in range(len(files2process)):
    f = open(os.path.join(input_directory_invoices, files2process[i]), "r", encoding='utf16')
    text = f.read()

    # PARSE ALL FIELDS
    peaje_acceso = ParseFields('Peaje de acceso: *(.+?) *Potencia', field_name = "peaje_acceso")
    numero_factura = ParseFields('Nº factura: *(.+?) *Periodo', field_name = "numero_factura")
    fecha_desde_periodo_consumo = ParseFields('Periodo de consumo *(.+?) *a[^\S\n\t]', field_name = "fecha_desde_periodo_consumo")
    fecha_hasta_periodo_consumo = ParseFields(re.escape(fecha_desde_periodo_consumo), ' *a *(.+?) *Fecha cargo', field_name = "fecha_hasta_periodo_consumo")
    fecha_cargo = ParseFields('límite de pago *(.+?) *Factura', field_name = "fecha_cargo")
    total_importe_factura = ParseFields('TOTAL IMPORTE FACTURA: *(.+?) *€', field_name = "total_importe_factura")
    numero_contador = ParseFields('Contador nº: *(.+?) *Lectura', field_name = "numero_contador")
    fecha_lectura_anterior = ParseFields('Lectura anterior real *(.+?) ', field_name = "fecha_lectura_anterior")
    lectura_anterior_periodo_punta = ParseFields(re.escape(fecha_lectura_anterior), ' *(.+?) *kWh', field_name = "lectura_anterior_periodo_punta")

    if peaje_acceso == '2.0DHA':
        lectura_anterior_periodo_valle = ParseFields(re.escape(fecha_lectura_anterior), ' *', re.escape(lectura_anterior_periodo_punta), ' *kWh *(.+?) *kWh', field_name = "lectura_anterior_periodo_valle")
    else:
        lectura_anterior_periodo_valle = ""

    fecha_lectura_actual = ParseFields('Lectura actual real *(.+?) ', field_name = "fecha_lectura_actual")
    lectura_actual_periodo_punta = ParseFields(re.escape(fecha_lectura_actual), ' *(.+?) *kWh', field_name = "lectura_actual_periodo_punta")

    if peaje_acceso == '2.0DHA':
        lectura_actual_periodo_valle = ParseFields(re.escape(fecha_lectura_actual), ' *', re.escape(lectura_actual_periodo_punta), ' *kWh *(.+?) *kWh', field_name = "lectura_actual_periodo_valle")
    else:
        lectura_actual_periodo_valle = ""

    consumo_periodo_punta = ParseFields('Consumo en el período \* *(.+?) *kWh', field_name = "consumo_periodo_punta")

    if peaje_acceso == '2.0DHA':
        consumo_periodo_valle = ParseFields('Consumo en el período \* *', re.escape(consumo_periodo_punta), ' *kWh *(.+?) *kWh', field_name = "consumo_periodo_valle")
    else:
        consumo_periodo_valle = ""

    # Sometimes there is no maxímetro data
    fecha_lectura_maximetro = ""
    lectura_real_maximetro = "0"
    fecha_lectura_maximetro = ParseFields('Lectura real maxímetro *(.+?) ', field_name = "fecha_lectura_maximetro")

    if not fecha_lectura_maximetro:
        print("fecha_lectura_maximetro -----> 0")
        print("lectura_real_maximetro -----> 0")
    else:
        lectura_real_maximetro = ParseFields('Lectura real maxímetro *', re.escape(fecha_lectura_maximetro), ' *(.+?) *kW', field_name = "lectura_real_maximetro")

    titular = ParseFields('Titular: *(.+?) *NIF', field_name = "titular")
    nif = ParseFields('NIF: *(.+?) *Dirección', field_name = "nif")
    direccion_suministro = ParseFields('Dirección suministro: *(.+?) *TIPO', field_name = "direccion_suministro")
    tipo_contrato = ParseFields('TIPO DE CONTRATO: *(.+?) *TIPO', field_name = "tipo_contrato")
    tipo_contador = ParseFields('TIPO DE CONTADOR: *(.+?) *Nº', field_name = "tipo_contador")
    numero_referencia = ParseFields('Nº referencia: *(.+?) *Oficina', field_name = "numero_referencia")
    potencia_contratada = ParseFields('Potencia contratada: *(.+?) *kW Referencia', field_name = "potencia_contratada")
    referencia_contrato_suministro = ParseFields('Referencia del contrato de suministro \(Comercializadora Regulada, Gas & Power, S.A.\): *(.+?) *Referencia', field_name = "referencia_contrato_suministro")
    referencia_contrato_acceso = ParseFields('Referencia del contrato de acceso \(UFD Distribución Electricidad, S.A.\): *(.+?) *Fecha', field_name = "referencia_contrato_acceso")
    fecha_final_contrato = ParseFields('Fecha final contrato: *(.+?) *\(renovación anual automática\)', field_name = "fecha_final_contrato")
    fecha_emision_factura = ParseFields('Fecha emisión factura: *(.+?) *Código', field_name = "fecha_emision_factura")
    cups = ParseFields('CUPS: *(.+?) *Atención', field_name = "cups")
    potencia_facturada_peaje_acceso = ParseFields('Importe por peaje de acceso: *(.+?) *kW', field_name = "potencia_facturada_peaje_acceso")
    precio_potencia_peaje_acceso = ParseFields('Importe por peaje de acceso: *', re.escape(potencia_facturada_peaje_acceso),' *kW *\* *(.+?) *€/kW', field_name = "precio_potencia_peaje_acceso")
    dias_periodo_potencia_peaje_acceso = ParseFields('y año *\* *\((.+?)/', field_name = "dias_periodo_potencia_peaje_acceso")
    dias_año_potencia_peaje_acceso = ParseFields('y año *\* *\(', re.escape(dias_periodo_potencia_peaje_acceso),'/(.+?)\)', field_name = "dias_año_potencia_peaje_acceso")
    importe_potencia_peaje_acceso = ParseFields('_+ *(.+?) *€ Importe por margen', field_name = "importe_potencia_peaje_acceso")
    potencia_facturada_margen_comercializadora = ParseFields('Importe por margen de comercialización fijo: *(.+?) *kW', field_name = "potencia_facturada_margen_comercializadora")
    precio_potencia_margen_comercializadora = ParseFields('Importe por margen de comercialización fijo: *', re.escape(potencia_facturada_margen_comercializadora), ' *kW *\* *(.+?) *€', field_name = "precio_potencia_margen_comercializadora")
    dias_periodo_margen_comercializadora = ParseFields(re.escape(potencia_facturada_margen_comercializadora), ' *kW *\* *', re.escape(precio_potencia_margen_comercializadora), ' *€/kW y año \* \((.+?)/', field_name = "dias_periodo_margen_comercializadora")
    dias_año_margen_comercializadora = ParseFields(re.escape(potencia_facturada_margen_comercializadora), ' *kW *\* *', re.escape(precio_potencia_margen_comercializadora), ' *€/kW y año \* \(', re.escape(dias_periodo_margen_comercializadora), '/(.+?)\)', field_name = "dias_año_margen_comercializadora")
    importe_margen_comercializadora_fijo = ParseFields('_+ *(\d{0,3},\d{0,3}?) *€ Facturación por energía', field_name = "importe_margen_comercializadora_fijo")

    if peaje_acceso == '2.0DHA':
        energia_consumida_periodo_punta = ParseFields('Importe por peaje de acceso punta: *(.+?) *kWh', field_name = "energia_consumida_periodo_punta")
        precio_energia_peaje_acceso_punta = ParseFields('Importe por peaje de acceso punta: *', re.escape(energia_consumida_periodo_punta), ' *kWh \(real\) \* (.+?) *€', field_name = "precio_energia_peaje_acceso_punta")
        importe_peaje_acceso_punta = ParseFields('_+ *(\d{0,3},\d{0,3}?) *€ Importe por coste de la energía punta', field_name = "importe_peaje_acceso_punta")
        precio_energia_punta_coste = ParseFields('Importe por coste de la energía punta: *', re.escape(energia_consumida_periodo_punta), ' *kWh[^\S\n\t]\(real\) *\* *(.+?) *€', field_name = "precio_energia_punta_coste")
        importe_energia_punta_coste = ParseFields('_+ *(\d{0,3},\d{0,3}?) *€ Importe por peaje de acceso valle', field_name = "importe_energia_punta_coste")
        energia_consumida_periodo_valle = ParseFields('Importe por peaje de acceso valle: *(.+?) *kWh', field_name = "energia_consumida_periodo_valle")
        precio_energia_peaje_acceso_valle = ParseFields('Importe por peaje de acceso valle: *', re.escape(energia_consumida_periodo_valle), ' *kWh *\(real\) *\* *(.+?) *€', field_name = "precio_energia_peaje_acceso_valle")
        importe_peaje_acceso_valle = ParseFields('_+ *(\d{0,3},\d{0,3}?) *€ Importe por coste de la energía valle', field_name = "importe_peaje_acceso_valle")
        precio_energia_valle_coste = ParseFields('Importe por coste de la energía valle: *', re.escape(energia_consumida_periodo_valle), ' *kWh *\(real\) *\* *(.+?) *€', field_name = "precio_energia_valle_coste")
        importe_energia_valle_coste = ParseFields('Importe por coste de la energía valle: *', re.escape(energia_consumida_periodo_valle), ' *kWh *\(real\) *\* *', re.escape(precio_energia_valle_coste), ' *€/kWh *', '_+ *(\d{0,3},\d{0,3}?) *€', field_name = "importe_energia_valle_coste")
    else:
        energia_consumida_periodo = ParseFields('del PVPC\). Importe por peaje de acceso: *(.+?) *kWh \(real\)', field_name = "energia_consumida_periodo")
        precio_energia_peaje_acceso = ParseFields('del PVPC\). Importe por peaje de acceso: *', re.escape(energia_consumida_periodo), ' *kWh \(real\) \* (.+?) *€', field_name = "precio_energia_peaje_acceso")
        importe_peaje_acceso = ParseFields('_+ *(\d{0,3},\d{0,3}?) *€ Importe por coste de la energía', field_name = "importe_peaje_acceso")
        precio_energia_coste = ParseFields('Importe por coste de la energía: *', re.escape(energia_consumida_periodo), ' *kWh[^\S\n\t]\(real\) *\* *(.+?) *€', field_name = "precio_energia_coste")
        importe_energia_coste = ParseFields('Importe por coste de la energía: *', re.escape(energia_consumida_periodo), ' *kWh[^\S\n\t]\(real\) *\* *', re.escape(precio_energia_coste), ' *€/kWh *_+ *(\d{0,3},\d{0,3}?) *€', field_name = "importe_energia_coste")

    suplemento_territorial = ParseFields('según la orden TEC/271/2019 \(\*\) *_+', ' *(\d{0,3},\d{0,3}?) *€ Subtotal', field_name = "suplemento_territorial")

    if not suplemento_territorial:
        suplemento_territorial = "0"

    base_impuesto_electrico = ParseFields('Impuesto Eléctrico *\((.+?) *\*', field_name = "base_impuesto_electrico")
    porcentaje_impuesto_electrico = ParseFields('Impuesto especial al tipo del *(.+?)%', field_name = "porcentaje_impuesto_electrico")
    importe_impuesto_electrico = ParseFields('_+ *(\d{0,3},\d{0,3}?) *€ *Alquiler equipos de medida y control', field_name = "importe_impuesto_electrico")
    dias_alquiler_equipos_medida_control = ParseFields('Alquiler de equipos de medida y control *\((.+?) *días', field_name = "dias_alquiler_equipos_medida_control")
    precio_alquiler_contador = ParseFields('Alquiler de equipos de medida y control *\(', re.escape(dias_alquiler_equipos_medida_control),' *días *\* *(.+?) *€/Día', field_name = "precio_alquiler_contador")
    importe_alquiler_contador = ParseFields('€/Día\) *_+ *(.+?) *€', field_name = "importe_alquiler_contador")
    porcentaje_iva = ParseFields('Impuesto I.V.A. al tipo del *(.+?)%', field_name = "porcentaje_iva")
    importe_iva = ParseFields('_+ *(\d{0,3},\d{0,3}?) *€ *TOTAL *IMPORTE', field_name = "importe_iva")

    # CHECKS AND ALERTS
    check_deviation_product(total = importe_potencia_peaje_acceso , w = potencia_facturada_peaje_acceso, x = precio_potencia_peaje_acceso, dias_periodo = dias_periodo_potencia_peaje_acceso, dias_año = dias_año_potencia_peaje_acceso)
    check_deviation_product(total = importe_margen_comercializadora_fijo , w = potencia_facturada_margen_comercializadora, x = precio_potencia_margen_comercializadora, dias_periodo = dias_periodo_margen_comercializadora, dias_año = dias_año_margen_comercializadora)
    check_deviation_sum(total = total_importe_factura, t = lectura_actual_periodo_punta, u = "-" + lectura_anterior_periodo_punta)

    if peaje_acceso == '2.0DHA':
        check_deviation_product(total = importe_energia_punta_coste , w = energia_consumida_periodo_punta, x = precio_energia_punta_coste)
        check_deviation_product(total = importe_peaje_acceso_valle , w = energia_consumida_periodo_valle, x = precio_energia_peaje_acceso_valle)
        check_deviation_product(total = importe_energia_valle_coste , w = energia_consumida_periodo_valle, x = precio_energia_valle_coste)
        check_deviation_sum(total = base_impuesto_electrico, t = importe_potencia_peaje_acceso, u = importe_margen_comercializadora_fijo, v = importe_peaje_acceso_punta, w = importe_energia_punta_coste, x = importe_peaje_acceso_valle, y = importe_energia_valle_coste, z = suplemento_territorial)

    else:
        check_deviation_product(total = importe_peaje_acceso , w = energia_consumida_periodo, x = precio_energia_peaje_acceso)
        check_deviation_product(total = importe_energia_coste , w = energia_consumida_periodo, x = precio_energia_coste)
        check_deviation_sum(total = base_impuesto_electrico, t = importe_potencia_peaje_acceso, u = importe_margen_comercializadora_fijo, v = importe_peaje_acceso, w = importe_energia_coste, x = suplemento_territorial)

    check_deviation_product(total = importe_impuesto_electrico , w = base_impuesto_electrico, x = porcentaje_impuesto_electrico, dias_periodo = "0,01")
    check_deviation_product(total = importe_alquiler_contador , w = dias_alquiler_equipos_medida_control, x = precio_alquiler_contador)
    check_deviation_sum(total = total_importe_factura, t = base_impuesto_electrico, u = importe_impuesto_electrico, v = importe_alquiler_contador, w = importe_iva)

    check_deviation_billed_power(float(lectura_real_maximetro.replace('.','DOT').replace('DOT','').replace(',','.')), float(potencia_contratada.replace('.','DOT').replace('DOT','').replace(',','.')), float(potencia_facturada_peaje_acceso.replace('.','DOT').replace('DOT','').replace(',','.')), float(potencia_facturada_margen_comercializadora.replace('.','DOT').replace('DOT','').replace(',','.')))
    # ROUND PRICES
    precio_potencia_peaje_acceso_2db = str(round_down(precio_potencia_peaje_acceso))
    precio_potencia_margen_comercializadora_2db = str(round_down(precio_potencia_margen_comercializadora))

    # PREPARE DATA TO EXPORT TO DATABASE
    peaje_acceso_2db = str(peaje_acceso)
    numero_factura_2db = str(numero_factura)
    fecha_desde_periodo_consumo_2db = str(fecha_desde_periodo_consumo)
    fecha_hasta_periodo_consumo_2db = str(fecha_hasta_periodo_consumo)
    fecha_cargo_2db = str(fecha_cargo)
    total_importe_factura_2db = str(total_importe_factura)
    numero_contador_2db = str(numero_contador)
    fecha_lectura_anterior_2db = str(fecha_lectura_anterior)
    lectura_anterior_periodo_punta_2db = str(lectura_anterior_periodo_punta)
    lectura_anterior_periodo_valle_2db = str(lectura_anterior_periodo_valle)
    fecha_lectura_actual_2db = str(fecha_lectura_actual)
    lectura_actual_periodo_punta_2db = str(lectura_actual_periodo_punta)
    lectura_actual_periodo_valle_2db = str(lectura_actual_periodo_valle)
    consumo_periodo_punta_2db = str(consumo_periodo_punta)
    consumo_periodo_valle_2db = str(consumo_periodo_valle)
    fecha_lectura_maximetro_2db = str(fecha_lectura_maximetro)
    lectura_real_maximetro_2db = str(lectura_real_maximetro)
    titular_2db = str(titular)
    nif_2db = str(nif)
    direccion_suministro_2db = str(direccion_suministro)
    tipo_contrato_2db = str(tipo_contrato)
    tipo_contador_2db = str(tipo_contador)
    numero_referencia_2db = str(numero_referencia)
    potencia_contratada_2db = str(potencia_contratada)
    referencia_contrato_suministro_2db = str(referencia_contrato_suministro)
    referencia_contrato_acceso_2db = str(referencia_contrato_acceso)
    fecha_final_contrato_2db = str(fecha_final_contrato)
    fecha_emision_factura_2db = str(fecha_emision_factura)
    cups_2db = str(cups)
    potencia_facturada_peaje_acceso_2db = str(potencia_facturada_peaje_acceso)
    precio_potencia_peaje_acceso_2db = str(precio_potencia_peaje_acceso)
    dias_periodo_potencia_peaje_acceso_2db = str(dias_periodo_potencia_peaje_acceso)
    dias_año_potencia_peaje_acceso_2db = str(dias_año_potencia_peaje_acceso)
    importe_potencia_peaje_acceso_2db = str(importe_potencia_peaje_acceso)
    importe_margen_comercializadora_fijo_2db = str(importe_margen_comercializadora_fijo)
    potencia_facturada_margen_comercializadora_2db = str(potencia_facturada_margen_comercializadora)
    precio_potencia_margen_comercializadora_2db = str(precio_potencia_margen_comercializadora)
    dias_periodo_margen_comercializadora_2db = str(dias_periodo_margen_comercializadora)
    dias_año_margen_comercializadora_2db = str(dias_año_margen_comercializadora)

    if peaje_acceso == '2.0DHA':
        energia_consumida_periodo_punta_2db = str(energia_consumida_periodo_punta)
        importe_peaje_acceso_punta_2db = str(importe_peaje_acceso_punta)
        importe_energia_punta_coste_2db = str(importe_energia_punta_coste)
        energia_consumida_periodo_valle_2db = str(energia_consumida_periodo_valle)
        importe_peaje_acceso_valle_2db = str(importe_peaje_acceso_valle)
        importe_energia_valle_coste_2db = str(importe_energia_valle_coste)
        precio_energia_peaje_acceso_punta_2db = str(round_down(precio_energia_peaje_acceso_punta))
        precio_energia_punta_coste_2db = str(round_down(precio_energia_punta_coste))
        precio_energia_peaje_acceso_valle_2db = str(precio_energia_peaje_acceso_valle)
        precio_energia_valle_coste_2db = str(round_down(precio_energia_valle_coste))

    else:
        precio_energia_peaje_acceso_punta_2db = str(round_down(precio_energia_peaje_acceso))
        importe_peaje_acceso_punta_2db = str(importe_peaje_acceso)
        importe_energia_punta_coste_2db = str(importe_energia_coste)
        precio_energia_punta_coste_2db = str(round_down(precio_energia_coste))
        energia_consumida_periodo_punta_2db = str(energia_consumida_periodo)
        energia_consumida_periodo_valle_2db = str("")
        precio_energia_peaje_acceso_valle_2db = str("")
        importe_peaje_acceso_valle_2db = str("")
        precio_energia_valle_coste_2db = str("")
        importe_energia_valle_coste_2db = str("")


    suplemento_territorial_2db = str(suplemento_territorial)
    base_impuesto_electrico_2db = str(base_impuesto_electrico)
    porcentaje_impuesto_electrico_2db = str(porcentaje_impuesto_electrico)
    importe_impuesto_electrico_2db = str(importe_impuesto_electrico)
    dias_alquiler_equipos_medida_control_2db = str(dias_alquiler_equipos_medida_control)
    precio_alquiler_contador_2db = str( precio_alquiler_contador)
    porcentaje_iva_2db = str(porcentaje_iva)
    importe_iva_2db = str(importe_iva)
    importe_alquiler_contador_2db = str(round_down(importe_alquiler_contador))
    importe_iva_2db = str(round_down(importe_iva))
    precio_alquiler_contador_2db = str(round_down(precio_alquiler_contador))

    mycursor = mydb.cursor()

    sql = "INSERT INTO m_i_ayuntamientos (titular, nif, direccion_suministro) VALUES (%s, %s, %s)"
    print(sql)
    val = [titular_2db, nif_2db, direccion_suministro_2db]
    print(val)
    mycursor.execute(sql, val)

    mydb.commit()

    print(mycursor.rowcount, "record inserted.")

    mycursor = mydb.cursor()

    sql = "INSERT INTO m_i_suministros (cups, peaje_acceso, numero_contador, tipo_contador, potencia_contratada, tipo_contrato, fecha_final_contrato, numero_referencia, referencia_contrato_suministro, referencia_contrato_acceso) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    print(sql)
    val = [cups_2db, peaje_acceso_2db, numero_contador_2db, tipo_contador_2db, potencia_contratada_2db, tipo_contrato_2db, fecha_final_contrato_2db, numero_referencia_2db, referencia_contrato_suministro_2db, referencia_contrato_acceso_2db]
    print(val)
    mycursor.execute(sql, val)

    mydb.commit()

    print(mycursor.rowcount, "record inserted.")

    mycursor = mydb.cursor()

    sql = "INSERT INTO p_i_facturas (numero_factura, cups, peaje_acceso, numero_contador, titular, fecha_emision_factura, fecha_desde_periodo_consumo, fecha_hasta_periodo_consumo, fecha_cargo, total_importe_factura, fecha_lectura_anterior, lectura_anterior_periodo_punta, lectura_anterior_periodo_valle, fecha_lectura_actual, lectura_actual_periodo_punta, lectura_actual_periodo_valle, consumo_periodo_punta, consumo_periodo_valle, fecha_lectura_maximetro, lectura_real_maximetro, potencia_contratada, potencia_facturada_peaje_acceso, precio_potencia_peaje_acceso, dias_periodo_potencia_peaje_acceso, dias_año_potencia_peaje_acceso, importe_potencia_peaje_acceso, potencia_facturada_margen_comercializadora, precio_potencia_margen_comercializadora, dias_periodo_margen_comercializadora, dias_año_margen_comercializadora, importe_margen_comercializadora_fijo, energia_consumida_periodo_punta, precio_energia_peaje_acceso_punta, importe_peaje_acceso_punta, precio_energia_punta_coste, importe_energia_punta_coste, energia_consumida_periodo_valle, precio_energia_peaje_acceso_valle, importe_peaje_acceso_valle, precio_energia_valle_coste, importe_energia_valle_coste, suplemento_territorial, base_impuesto_electrico, porcentaje_impuesto_electrico, importe_impuesto_electrico, dias_alquiler_equipos_medida_control, precio_alquiler_contador, importe_alquiler_contador, porcentaje_iva, importe_iva) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    print(sql)
    val = [numero_factura_2db, cups_2db, peaje_acceso_2db, numero_contador_2db, titular_2db, fecha_emision_factura_2db, fecha_desde_periodo_consumo_2db, fecha_hasta_periodo_consumo_2db, fecha_cargo_2db, total_importe_factura_2db, fecha_lectura_anterior_2db, lectura_anterior_periodo_punta_2db, lectura_anterior_periodo_valle_2db, fecha_lectura_actual_2db, lectura_actual_periodo_punta_2db, lectura_actual_periodo_valle_2db, consumo_periodo_punta_2db, consumo_periodo_valle_2db, fecha_lectura_maximetro_2db, lectura_real_maximetro_2db, potencia_contratada_2db, potencia_facturada_peaje_acceso_2db, precio_potencia_peaje_acceso_2db, dias_periodo_potencia_peaje_acceso_2db, dias_año_potencia_peaje_acceso_2db, importe_potencia_peaje_acceso_2db, potencia_facturada_margen_comercializadora_2db, precio_potencia_margen_comercializadora_2db, dias_periodo_margen_comercializadora_2db, dias_año_margen_comercializadora_2db, importe_margen_comercializadora_fijo_2db, energia_consumida_periodo_punta_2db, precio_energia_peaje_acceso_punta_2db, importe_peaje_acceso_punta_2db, precio_energia_punta_coste_2db, importe_energia_punta_coste_2db, energia_consumida_periodo_valle_2db, precio_energia_peaje_acceso_valle_2db, importe_peaje_acceso_valle_2db, precio_energia_valle_coste_2db, importe_energia_valle_coste_2db, suplemento_territorial_2db, base_impuesto_electrico_2db, porcentaje_impuesto_electrico_2db, importe_impuesto_electrico_2db, dias_alquiler_equipos_medida_control_2db, precio_alquiler_contador_2db, importe_alquiler_contador_2db, porcentaje_iva_2db, importe_iva_2db]
    print(val)
    mycursor.execute(sql, val)

    mydb.commit()

    print(mycursor.rowcount, "record inserted.")