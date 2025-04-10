import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from fpdf import FPDF
import win32com.client as win32
from datetime import datetime
import numpy as np
import re
import unicodedata
import camelot
import requests

import pdfplumber



# ======================== CONFIGURACIÓN ========================
ruta_carpeta = r"C:\Users\jabar\OneDrive\Escritorio\PruebaFiltro"
df_filtrado = None  # Variable global para almacenar los datos filtrados

# ======================== FUNCIONES ========================
def obtener_ultimo_excel():
    nombre_archivo = "base.xlsx"  # <- Cambialo por el nombre real del archivo
    ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)
    return ruta_archivo if os.path.exists(ruta_archivo) else None
   # archivos = glob.glob(os.path.join(ruta_carpeta, "*.xlsx"))
   #return max(archivos, key=os.path.getctime) if archivos else None

def procesar_segundo_excel():
    archivo_adicional = os.path.join(ruta_carpeta, "ExportacaoAE_DNT-Ciclo 3.xlsx")  # Cambia por el nombre real

    if not os.path.exists(archivo_adicional):
        messagebox.showerror("Error", "No se encontró el segundo archivo Excel.")
        return

    df_extra = pd.read_excel(archivo_adicional, usecols=["SECTOR", "PERÍODO"])

    # Limpieza de columnas
    df_extra["SECTOR"] = df_extra["SECTOR"].astype(str).str.strip()
    df_extra["PERÍODO"] = df_extra["PERÍODO"].astype(str)

    # Mapeo de SECTOR a nombres
    mapeo_nombres = {
        "50602": "ANA KAREN MORA",
        "50603": "RICARDO ZUNIGA",
        "50604": "KAROLINA MONTERO",
        "50605": "KEREN CARVAJAL",
        "50606": "JOHAN ALBERTO ARCE NÚÑEZ"
        # Agrega más según lo que necesites
    }
    df_extra["NOMBRE_SECTOR"] = df_extra["SECTOR"].map(mapeo_nombres)
  
   
    # Extraer número de horas del texto como "8 horas"
    df_extra["HORAS"] = df_extra["PERÍODO"].apply(lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)

    # Suma total de horas por nombre
    resumen_horas = df_extra.groupby("NOMBRE_SECTOR")["HORAS"].sum().reset_index()

    print("Resumen de horas por nombre:")
    print(resumen_horas)

    # Si querés también la suma total de todo:
    total_horas = df_extra["HORAS"].sum()
    print(f"Total de horas (todos los sectores): {total_horas}")

    return resumen_horas, total_horas

def aplicar_filtros():
  

    CICLOS = {
        "CICLO 1": (datetime(2025, 1, 1), datetime(2025, 2, 10)),
        "CICLO 2": (datetime(2025, 2, 11), datetime(2025, 3, 13)),
        "CICLO 3": (datetime(2025, 3, 14), datetime(2025, 4, 12)),
        "CICLO 4": (datetime(2025, 4, 13), datetime(2025, 5, 15)),
        # Agregá más ciclos según sea necesario
    }

    def obtener_ciclo_actual(fecha_actual):
        for ciclo, (inicio, fin) in CICLOS.items():
            if inicio <= fecha_actual <= fin:
                return ciclo, inicio, fin
        return None, None, None
    fecha_actual = datetime.now()
    ciclo_actual, fecha_ciclo_inicio, fecha_ciclo_fin = obtener_ciclo_actual(fecha_actual)
    if not ciclo_actual:
        messagebox.showerror("Error", "No se pudo determinar el ciclo actual basado en la fecha.")
        return

    print(f"Ciclo actual: {ciclo_actual} ({fecha_ciclo_inicio.date()} a {fecha_ciclo_fin.date()})")

    global df_filtrado
    archivo_excel = obtener_ultimo_excel()
    if not archivo_excel:
        messagebox.showerror("Error", "No se encontró ningún archivo Excel en la carpeta.")
        return

    df = pd.read_excel(archivo_excel)

    columnas_necesarias = {"SECTOR", "NOMBRE DEL SECTOR", "FECHA DE VISITA", "CLASIFICACIÓN", "TIPO CLIENTE"}
    if not columnas_necesarias.issubset(df.columns):
        messagebox.showerror("Error", f"El archivo Excel no tiene las columnas necesarias: {columnas_necesarias}")
        return

    df["FECHA DE VISITA"] = pd.to_datetime(df["FECHA DE VISITA"], errors="coerce", dayfirst=True)

    nombre_filtro = entrada_nombre.get().strip()
    fecha_inicio = entrada_fecha_inicio.get().strip()
    fecha_fin = entrada_fecha_fin.get().strip()
    pais_seleccionado = combo_pais.get()
    tipo_cliente_seleccionado = combo_tipo_cliente.get()
    clasificacion_seleccionada = [clasificacion_listbox.get(i) for i in clasificacion_listbox.curselection()]
    global nom_correo
    nom_correo = nombre_filtro
    if nombre_filtro:
        df = df[df["NOMBRE DEL SECTOR"].str.contains(nombre_filtro, case=False, na=False)]
    if fecha_inicio:
        df = df[df["FECHA DE VISITA"] >= pd.to_datetime(fecha_inicio, errors="coerce", dayfirst=True)]
    if fecha_fin:
        df = df[df["FECHA DE VISITA"] <= pd.to_datetime(fecha_fin, errors="coerce", dayfirst=True)]
    if clasificacion_seleccionada:
        df = df[df["CLASIFICACIÓN"].isin(clasificacion_seleccionada)]
    
    if pais_seleccionado == "Costa Rica":
        df = df[df["SECTOR"].astype(str).str.startswith("506")]
    elif pais_seleccionado == "Panamá":
        df = df[df["SECTOR"].astype(str).str.startswith("507")]
    elif pais_seleccionado == "Nicaragua":
        df = df[df["SECTOR"].astype(str).str.startswith("505")]
    elif pais_seleccionado == "Salvador":
        df = df[df["SECTOR"].astype(str).str.startswith("503")]
    elif pais_seleccionado == "Honduras":
        df = df[df["SECTOR"].astype(str).str.startswith("504")]
    elif pais_seleccionado == "Guatemala":
        df = df[df["SECTOR"].astype(str).str.startswith("502")]    

    if tipo_cliente_seleccionado != "Todos":
        df = df[df["TIPO CLIENTE"] == tipo_cliente_seleccionado]

    df_filtrado = df.copy()
    for row in tabla.get_children():
        tabla.delete(row)

    if df.empty:
        messagebox.showwarning("Sin Resultados", "No se encontraron registros con esos filtros.")
    else:
        for _, fila in df.iterrows():
            tabla.insert("", "end", values=(fila["SECTOR"], fila["NOMBRE DEL SECTOR"], fila["FECHA DE VISITA"].strftime('%d/%m/%Y'), fila["CLASIFICACIÓN"], fila["TIPO CLIENTE"]))
            # Agregar una fila de total al final de la tabla
        global total_registros
        total_registros = len(df)
        tabla.insert("", "end", values=("TOTAL", "", "", total_registros, ""))
#///////////////////////////////////////////////////DATOS PARA SACAR LA RESTA DE DIAS AL CICLO///////////////////////////////////////////////////////////////////////////////////
    archivo_adicional = obtener_excel_especifico("ExportacaoAE_DNT-Ciclo 3.xlsx")
    if not archivo_adicional:
        return

    df_extra = pd.read_excel(archivo_adicional, usecols=["SECTOR CLIENTE", "PERÍODO", "FECHA_INCLUSION"])
    df_extra.columns = df_extra.columns.str.strip()
    df_extra["SECTOR CLIENTE"] = df_extra["SECTOR CLIENTE"].astype(str).str.strip()
    df_extra["PERÍODO"] = df_extra["PERÍODO"].astype(str)
    df_extra["FECHA_INCLUSION"] = pd.to_datetime(df_extra["FECHA_INCLUSION"], errors="coerce", dayfirst=True)
    df_extra = df_extra[
        (df_extra["FECHA_INCLUSION"] >= fecha_ciclo_inicio) &
        (df_extra["FECHA_INCLUSION"] <= fecha_ciclo_fin)
    ]  


    mapeo_nombres = {
        "50602": "ANA KAREN MORA",
        "50603": "RICARDO  ZUNIGA",
        "50604": "KAROLINA MONTERO",
        "50605": "KEREN CARVAJAL",
        "50606": "JOHAN ALBERTO ARCE NÚÑEZ"
    }

    # Buscar el nombre completo exacto filtrado en el primer Excel
    nombre_filtrado_completo = None
    if not df_filtrado.empty:
        nombre_filtrado_completo = df_filtrado["NOMBRE DEL SECTOR"].iloc[0].strip().upper()
    def normalizar(texto):
        return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8').strip().upper().replace("  ", " ")

    # Encontrar el código de sector basado en ese nombre
    codigo_sector = None
    if nombre_filtrado_completo:
        nombre_normalizado = normalizar(nombre_filtrado_completo)
        for codigo, nombre in mapeo_nombres.items():
            if nombre_normalizado == normalizar(nombre):
                codigo_sector = codigo
                break

    if not codigo_sector:
        print(f"No se encontró código para el nombre '{nombre_filtrado_completo}'.")
        return

    # === Filtrar el segundo Excel por el código encontrado ===
    df_extra_filtrado = df_extra[df_extra["SECTOR CLIENTE"] == codigo_sector]

    # === Extraer y sumar horas ===
    df_extra_filtrado.loc[:, "HORAS"] = df_extra_filtrado["PERÍODO"].apply(
        lambda x: int(re.search(r'\d+', x).group()) if pd.notnull(x) and re.search(r'\d+', x) else 0
    )

    # Filtrar horas mayores o iguales a 6
    df_extra_filtrado = df_extra_filtrado[df_extra_filtrado["HORAS"] >= 6]

    global RESTAR_HORAS
    RESTAR_HORAS = df_extra_filtrado["HORAS"].sum()

    print(f"Total de horas para {nombre_filtrado_completo} (código {codigo_sector}): {RESTAR_HORAS}")


def obtener_excel_especifico(nombre_archivo):
    archivo_path = os.path.join(ruta_carpeta, nombre_archivo)
    return archivo_path if os.path.exists(archivo_path) else None




def generar_pdf():
    if df_filtrado is None or df_filtrado.empty:
        messagebox.showwarning("Advertencia", "No hay datos filtrados para generar el PDF.")
        return

    ruta_pdf = os.path.join(ruta_carpeta, "reporte_filtrado_individual.pdf")
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    # Agregar espacio antes de la tabla resumen
    pdf.ln(9)

    print("Existe la fuente?", os.path.exists("DejaVuSans.ttf"))  # Esto debe imprimir True

    

    # ✅ Cargar fuente compatible con emojis
    pdf.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
    pdf.add_font('DejaVu', 'B', 'DejaVuSans-Bold.ttf', uni=True)    # Bold
    pdf.set_font('DejaVu', '', 10)

    # Calcular los totales para la tabla resumen
    total_medicos = len(df_filtrado)
    total_vip = len(df_filtrado[df_filtrado["CLASIFICACIÓN"] == "VIP"])
    total_regentes = len(df_filtrado[df_filtrado["CLASIFICACIÓN"].isna()])
    var_arreglo = total_medicos - total_regentes 
    print(total_medicos) # BORRAR AHORITA ES UNA PRUEBA DE CONSOLA

    df_filtrado["FECHA DE VISITA"] = pd.to_datetime(df_filtrado["FECHA DE VISITA"])
    fechas_unicas = sorted(df_filtrado["FECHA DE VISITA"].dropna().unique())

    # Sumar los totales de la tabla detallada
    total_medicos_dias = sum(len(df_filtrado[df_filtrado["FECHA DE VISITA"] == fecha]) for fecha in fechas_unicas)
    total_vip_dias = sum(len(df_filtrado[(df_filtrado["FECHA DE VISITA"] == fecha) & (df_filtrado["CLASIFICACIÓN"] == "VIP")]) for fecha in fechas_unicas)
    total_regentes_dias = sum(len(df_filtrado[(df_filtrado["FECHA DE VISITA"] == fecha) & (df_filtrado["CLASIFICACIÓN"].isna())]) for fecha in fechas_unicas)
    # Evitar divisiones por cero
    global total_dias
    total_dias = len(fechas_unicas) if len(fechas_unicas) > 0 else 1  

    # Calcular Cobertura Diaria
    global cobertura_medicos# = round(var_arreglo/ total_dias , 2) # total_medicos_dias
    cobertura_medicos = round(var_arreglo/ total_dias , 2) # total_medicos_dias
    print(total_medicos)
    print(total_dias)
    global cobertura_vip
    cobertura_vip = round(total_vip_dias / total_dias, 2)
    global cobertura_regentes
    cobertura_regentes = round(total_regentes_dias / total_dias, 2)
    print(nom_correo)

    if nom_correo == "ANA KAREN MORA":
        divisor_medicos = 150
        divisor_vip = 36
        divisor_regentes = 80
        Var_arreglo_real = 7.5
        Var_arreglo_vip = 1.8
        Var_arreglo_estandar = 5.7
        Var_arreglo_regentes = 4
        print("ENTRO EL IF")
    elif nom_correo == "RICARDO ZUNIGA":
        divisor_medicos = 180
        divisor_vip = 40
        divisor_regentes = 80
        Var_arreglo_real = 9
        Var_arreglo_vip = 2.08
        Var_arreglo_vip = 6.2
        Var_arreglo_regentes = 4
        print("No entro al if")
    elif nom_correo == "KEREN CARVAJAL":
        divisor_medicos = 180
        divisor_vip = 40
        divisor_regentes = 80
        Var_arreglo_real = 9.1
        Var_arreglo_vip = 2.05
        Var_arreglo_estandar = 7.05
        Var_arreglo_regentes = 4
        print("No entro al if")
    elif nom_correo == "JOHAN ALBERTO ARCE NUÑEZ":
        divisor_medicos = 180
        divisor_vip = 40
        divisor_regentes = 80
        Var_arreglo_real = 8.9
        Var_arreglo_vip = 2.3
        Var_arreglo_estandar = 6.6
        Var_arreglo_regentes = 4
        print("No entro al if")


    # ACA LOS VOY A SETEAR COMO SI FUERA PARA ANA KAREN

    Conversion_horas_dias = RESTAR_HORAS / 8
    global Dias_Reales_Periodo
    Dias_Reales_Periodo = total_dias - Conversion_horas_dias
    var_arreglo_real = Var_arreglo_real * Dias_Reales_Periodo # CAMBIAR ESE VAR_ARREGLO POR EL TOTAL DE DIAS DE COBERTURA, O SEA EL TOTAL DE DIAS QUE JALA DE LA BASE 
   # print("aca total de dias de ciclo que llevamos", total_dias)
   # print("PRUEBA DE PRINT")
    #print(var_arreglo_real)
    var_arreglo_vip = Var_arreglo_vip * Dias_Reales_Periodo
    var_arreglo_regentes =  Var_arreglo_regentes * Dias_Reales_Periodo


    Cobertura_real_porcentaje = var_arreglo / var_arreglo_real
    global Co1
    Co1 = Cobertura_real_porcentaje * 100
    print(var_arreglo)
    print(var_arreglo_real)

    Cobertura_Real_vip = total_vip / var_arreglo_vip
    global Co2
    Co2 = Cobertura_Real_vip * 100
    #print(Cobertura_Real_vip)
    Cobertura_Real_regentes = total_regentes / var_arreglo_regentes
    global Co3
    Co3 = Cobertura_Real_regentes * 100

   
    pdf.set_font("DejaVu", "B", 10)
    pdf.cell(145, 10, "DATOS REALES", 1, ln=True, align="C")  
    pdf.set_font("DejaVu", "B", 10)
    pdf.cell(20, 10, "", 1)
    pdf.cell(50, 10, "Contactos Total", 1)
    pdf.cell(35, 10, "Cobertura", 1)
    pdf.cell(40, 10, "Promedio Diario", 1)

    pdf.ln()

    pdf.set_font("DejaVu", "", 10)
    pdf.cell(20, 10, "MÉDICOS", 1)
    pdf.cell(50, 10, str(var_arreglo), 1) #ACA ERA TOTAL_MEDICO
    if Co1 < 95:
        pdf.set_text_color(255, 0, 0)  # Rojo
        pdf.cell(35, 10, f"[RED]{Co1:.2f}%", 1)
        pdf.set_text_color(0, 0, 0)    # Restaurar a negro
    elif Co1 > 95 and Co1 < 99: 
        pdf.set_text_color(255, 255, 0)  # Amarillo
        pdf.cell(35, 10,f"[AMARILLO]{Co1:.2f}%", 1)
        pdf.set_text_color(0, 0, 0)
    else:
        pdf.set_text_color(255, 255, 0)  # verde
        pdf.cell(35, 10,f"[VERDE]{Co1:.2f}%", 1)
        pdf.set_text_color(0, 0, 0) 
    pdf.cell(40, 10, str(cobertura_medicos), 1)
    #valor1(cobertura_medicos)
    pdf.ln()


    pdf.cell(20, 10, "VIP", 1)
    pdf.cell(50, 10, str(total_vip), 1)
    if Co2 < 94.5:
        pdf.set_text_color(255, 0, 0)  # Rojo
        pdf.cell(35, 10, f"[RED]{Co2:.2f}%", 1)
        pdf.set_text_color(0, 0, 0)    # Restaurar a negro
    elif Co2 > 95 and Co2 < 99.5: 
        pdf.set_text_color(255, 255, 0)  # Amarillo
        pdf.cell(35, 10,f"[AMARILLO]{Co2:.2f}%", 1)
        pdf.set_text_color(0, 0, 0)
    else:
        pdf.set_text_color(255, 255, 0)  # verde
        pdf.cell(35, 10,f"[VERDE]{Co2:.2f}%", 1)
        pdf.set_text_color(0, 0, 0) 
    pdf.cell(40, 10, str(cobertura_vip), 1)
    pdf.ln()

    pdf.cell(20, 10, "REGENTES", 1)
    pdf.cell(50, 10, str(total_regentes), 1)
    if Co3 < 95:
        pdf.set_text_color(255, 0, 0)  # Rojo
        pdf.cell(35, 10, f"[RED]{Co3:.2f}%", 1)
        pdf.set_text_color(0, 0, 0)    # Restaurar a negro
    elif Co3 > 95 and Co3 < 99: 
        pdf.set_text_color(255, 255, 0)  # Amarillo
        pdf.cell(35, 10,f"[AMARILLO]{Co3:.2f}%", 1)
        pdf.set_text_color(0, 0, 0)
    else:
        pdf.set_text_color(255, 255, 0)  # verde
        pdf.cell(35, 10,f"[VERDE]{Co3:.2f}%", 1)
        pdf.set_text_color(0, 0, 0) 
    pdf.cell(40, 10,str(cobertura_regentes), 1)
    pdf.ln()

    # -------------------- REPORTE DETALLADO --------------------
    pdf.set_font("Arial", "B", 14)
    pdf.cell(270, 10, "Reporte Detallado por Día", ln=True, align="C")
    pdf.ln(10)

    df_filtrado["FECHA DE VISITA"] = pd.to_datetime(df_filtrado["FECHA DE VISITA"])
    fechas_unicas = sorted(df_filtrado["FECHA DE VISITA"].dropna().unique())

    # Función para verificar espacio antes de cada fila
    def check_page_space():
        if pdf.get_y() > 180:  # Si está cerca del final de la página
            pdf.add_page()
            agregar_encabezado()

    # Función para repetir el encabezado en cada nueva página
    def agregar_encabezado():
        pdf.set_font("Arial", "B", 12)
        pdf.cell(30, 10, "Clasificación", 1)
        for fecha in fechas_unicas:
            pdf.cell(15,10, fecha.strftime('%d/%m'), 1)  # Formato corto para ahorrar espacio
        pdf.ln()

    # Agregar encabezado inicial
    agregar_encabezado()

    # Fila "Médicos"
    #check_page_space()
    pdf.cell(30, 10, "MÉDICOS", 1)
    for fecha in fechas_unicas:
        df_dia = df_filtrado[df_filtrado["FECHA DE VISITA"] == fecha]
        total_medicos_dia = len(df_dia[df_dia["CLASIFICACIÓN"].isin(["VIP", "ESTANDAR"])])
        pdf.cell(15, 10, str(total_medicos_dia), 1)
    pdf.ln()

    # Fila "VIP"
   # check_page_space()
    pdf.cell(30, 10, "VIP", 1)
    for fecha in fechas_unicas:
        df_dia = df_filtrado[df_filtrado["FECHA DE VISITA"] == fecha]
        total_vip_dia = len(df_dia[df_dia["CLASIFICACIÓN"] == "VIP"])
        pdf.cell(15, 10, str(total_vip_dia), 1)
    pdf.ln()

    # Fila "Regentes"
    #check_page_space()
    pdf.cell(30, 10, "REGENTES", 1)
    for fecha in fechas_unicas:
        df_dia = df_filtrado[df_filtrado["FECHA DE VISITA"] == fecha]
        total_regentes_dia = len(df_dia[df_dia["CLASIFICACIÓN"].isna()])
        pdf.cell(15, 10, str(total_regentes_dia), 1)
    pdf.ln()



    # Guardar el PDF
    pdf.output(ruta_pdf)
    messagebox.showinfo("PDF Generado", f"El reporte se ha guardado en:\n{ruta_pdf}")
    return ruta_pdf



def valor1(valo1):
    print(valor1)
    
    
def extraer_todas_las_tablas_pdf(ruta_pdf):
    import pdfplumber

    tabla_html = ""
    primera_tabla_normal = True
    encabezados_columna = []

    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            tablas = page.extract_tables()

            if tablas:
                for tabla in tablas:
                    tabla_limpia = []

                    for fila in tabla:
                        if fila and any(celda and str(celda).strip() for celda in fila):
                            tabla_limpia.append([celda.strip() if celda else "" for celda in fila])

                    if not tabla_limpia:
                        continue

                    if (
                        len(tabla_limpia) >= 4 and
                        tabla_limpia[0][0].upper().startswith("DATOS REALES")
                    ):
                        tabla_html += """
                        <table border='1' style='border-collapse: collapse; width: 100%; text-align: center;'>
                            <tr>
                                <td colspan='4' style='font-weight: bold; padding: 8px;'>DATOS REALES</td>
                            </tr>
                        """

                        for fila in tabla_limpia[1:]:
                            if len(fila) >= 4:
                                tabla_html += "<tr>"
                                for celda in fila[:4]:
                                    estilo = "padding: 5px; border: 1px solid #000;"

                                    if "[RED]" in celda:
                                        estilo += "color: red; font-weight: bold;"
                                        celda = celda.replace("[RED]", "")
                                    elif "[AMARILLO]" in celda:
                                        estilo += "color: orange; font-weight: bold;"
                                        celda = celda.replace("[AMARILLO]", "")
                                    elif "[VERDE]" in celda:
                                        estilo += "color: green; font-weight: bold;"
                                        celda = celda.replace("[VERDE]", "")

                                    tabla_html += f"<td style='{estilo}'>{celda}</td>"
                                tabla_html += "</tr>"
                        tabla_html += "</table><br>"

                    else:
                        tabla_html += "<table border='1' style='border-collapse: collapse; width: 100%; text-align: center;'>"
                        for i, fila in enumerate(tabla_limpia):
                            tabla_html += "<tr>"
                            for j, celda in enumerate(fila):
                                estilo = "padding: 5px; border: 1px solid #000;"

                                # Marcadores de color
                                if "[RED]" in celda:
                                    estilo += "color: red; font-weight: bold;"
                                    celda = celda.replace("[RED]", "")
                                elif "[AMARILLO]" in celda:
                                    estilo += "color: orange; font-weight: bold;"
                                    celda = celda.replace("[AMARILLO]", "")
                                elif "[VERDE]" in celda:
                                    estilo += "color: green; font-weight: bold;"
                                    celda = celda.replace("[VERDE]", "")

                                tabla_html += f"<td style='{estilo}'>{celda}</td>"
                            tabla_html += "</tr>"
                        tabla_html += "</table><br>"

                        primera_tabla_normal = False

    return tabla_html if tabla_html else "<p>No se encontraron tablas en el PDF.</p>"


def es_correo_valido(correo):
    """Verifica si el correo tiene un formato válido usando una expresión regular."""
    patron = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(patron, correo) is not None

def enviar_correo():

    if nom_correo == "ANA KAREN MORA":
    #ruta_pdf = generar_pdf()
        correo_destino = entrada_correo.get().strip()
        if not correo_destino:
            messagebox.showwarning("Advertencia", "No se ingresó un correo. El reporte no se envió.")
            return
        
        if not es_correo_valido(correo_destino):
            messagebox.showerror("Error", "El correo ingresado no es válido. Inténtalo de nuevo.")
            return

        
        ruta_pdf = generar_pdf()


        #if not correo_destino:
        #    messagebox.showwarning("Advertencia", "No se ingresó un correo. El reporte no se envió.")
        #    return
        
        tabla_html = extraer_todas_las_tablas_pdf(ruta_pdf)  # Extrae todas las tablas como HTML
        Nombre = entrada_nombre.get().strip()
    # print(nom_correo)
        correo_cc = "jabarca@innoviahealthcare.com; jose03pif@gmail.com"  # Correos separados por punto y coma

        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = correo_destino
        mail.CC = correo_cc
        mail.Subject = "AGD - Requiere atencion: registro de tus visitas"
        mail.HTMLBody = f"""
        <p>Estimada {Nombre}:</p>
        <p>Aquí te comento el estado de los principales indicadores de tus visitas. También adjunto un cuadro con el registro diario de las mismas según surgen del fichero</p>
        <p>Día objetivo {total_dias}:</p> 
        <p>Día Reales {Dias_Reales_Periodo}:</p>
        {tabla_html}
        <b>Promedio de visitas a médicos = {cobertura_medicos}:</b>
        <p>El promedio mínimo para alcanzar los contactos programados es de 7 visitas diarias.</p>
        <b>Promedio de visitas a farmacias = {cobertura_regentes} :</b>
        <p>El promedio mínimo para alcanzar los contactos programados es de 4 visitas diarias.</p>
        <b>Cobertura médicos = {Co1:.2f}</b>
        <p>La cobertura teórica objetivo al día de la fecha es de 100.0%.</p>
        <b>Cobertura farmacias = {Co2:.2f}</b>
        <p>La cobertura teórica objetivo al día de la fecha es de 100.0%.</p>
        <b>Cobertura médicos VIP = {Co3:.2f}</b>
        <p>La cobertura teórica objetivo al día de la fecha es de 100.0%.</p>
        <p>Muchas gracias por tu atención!</p>


        """

        mail.Attachments.Add(ruta_pdf)

        mail.Send()
        messagebox.showinfo("Correo Enviado", "El correo ha sido enviado correctamente.")
    else:
        correo_destino = entrada_correo.get().strip()
        if not correo_destino:
            messagebox.showwarning("Advertencia", "No se ingresó un correo. El reporte no se envió.")
            return
        if not es_correo_valido(correo_destino):
            messagebox.showerror("Error", "El correo ingresado no es válido. Inténtalo de nuevo.")
            return
        
        ruta_pdf = generar_pdf()
        tabla_html = extraer_todas_las_tablas_pdf(ruta_pdf)  # Extrae todas las tablas como HTML
        Nombre = entrada_nombre.get().strip()
        correo_cc = "xotoyip215@buides.com; xotoyip215@buides.com"  # Correos separados por punto y coma

        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.CC = correo_cc
        mail.To = correo_destino
        mail.Subject = "AGD - Requiere atencion: registro de tus visitas"
        mail.HTMLBody = f"""
        <p>Estimada {Nombre}:</p>
        <p>Aquí te comento el estado de los principales indicadores de tus visitas. También adjunto un cuadro con el registro diario de las mismas según surgen del fichero</p>
        <p>Día objetivo {total_dias}:</p> 
        <p>Día Reales {Dias_Reales_Periodo}:</p>
        {tabla_html}
        <b>Promedio de visitas a médicos = {cobertura_medicos}:</b>
        <p>El promedio mínimo para alcanzar los contactos programados es de 7 visitas diarias.</p>
        <b>Promedio de visitas a farmacias = {cobertura_regentes} :</b>
        <p>El promedio mínimo para alcanzar los contactos programados es de 3 visitas diarias.</p>
        <b>Cobertura médicos = {Co1:.2f}</b>
        <p>La cobertura teórica objetivo al día de la fecha es de 100.0%.</p>
        <b>Cobertura farmacias = {Co2:.2f}</b>
        <p>La cobertura teórica objetivo al día de la fecha es de 100.0%.</p>
        <b>Cobertura médicos VIP = {Co3:.2f}</b>
        <p>La cobertura teórica objetivo al día de la fecha es de 100.0%.</p>
        <p>Muchas gracias por tu atención!</p>
        """
        mail.Attachments.Add(ruta_pdf)
        mail.Send()
        messagebox.showinfo("Correo Enviado", "El correo ha sido enviado correctamente.")
# ======================== VENTANAS ========================
def abrir_sistema_reportes():
    global root, tabla, entrada_nombre, entrada_fecha_inicio, entrada_fecha_fin, combo_pais, combo_tipo_cliente, clasificacion_listbox, entrada_correo

    root = tk.Toplevel()
    root.title("Sistema de Reportes")
    root.geometry("1000x600")

    
    nombres_consultores = ["ANA KAREN MORA", "RICARDO","KAROLINA MONTER", "KEREN CARVAJAL","FIORELLA CHACON","JOHAN ALBERTO",
    "LOURDES ROA", "NAZARETH NAVARRO","FANNY ROJAS", "MARIA GABRIELA DUARTE", "GENESIS TAMARA FUENTES BOJORG"]

    # Etiquetas y Entradas
    ttk.Label(root, text="Consultores:").grid(row=0, column=0, padx=7, pady=3)
    entrada_nombre = ttk.Combobox(root, values=nombres_consultores, state="readonly")
    entrada_nombre.grid(row=0, column=1, padx=10, pady=3)

    ttk.Label(root, text="Fecha Inicio (dd/mm/yyyy):").grid(row=1, column=0, padx=10, pady=5)
    entrada_fecha_inicio = ttk.Entry(root)
    entrada_fecha_inicio.grid(row=1, column=1, padx=10, pady=5)

    ttk.Label(root, text="Fecha Fin (dd/mm/yyyy):").grid(row=2, column=0, padx=10, pady=5)
    entrada_fecha_fin = ttk.Entry(root)
    entrada_fecha_fin.grid(row=2, column=1, padx=10, pady=5)

    # Filtro de país (nuevo)
    ttk.Label(root, text="País:").grid(row=3, column=0, padx=10, pady=5)
    combo_pais = ttk.Combobox(root, values=["Todos", "Costa Rica", "Panamá", "Nicaragua", "Salvador","Honduras","Guatemala"])
    combo_pais.grid(row=3, column=1, padx=10, pady=5)
    combo_pais.current(0)  # Seleccionar "Todos" por defecto

    # Filtro de TIPO CLIENTE
    ttk.Label(root, text="Tipo Cliente:").grid(row=4, column=0, padx=10, pady=5)
    combo_tipo_cliente = ttk.Combobox(root, values=["Todos", "PDV", "PROFISSIONAL"])
    combo_tipo_cliente.grid(row=4, column=1, padx=10, pady=5)
    combo_tipo_cliente.current(0)  # Seleccionar "Todos" por defecto

    # Filtro para clasificación
    ttk.Label(root, text="Clasificación:").grid(row=5, column=0, padx=10, pady=5)
    clasificacion_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, height=4)
    clasificacion_listbox.insert(tk.END, "VIP", "GENERAL", "REGENTES")
    clasificacion_listbox.grid(row=5, column=1, padx=10, pady=5)

    # Botón para aplicar filtros
    btn_filtrar = ttk.Button(root, text="🔍 Aplicar Filtros", command=aplicar_filtros)
    btn_filtrar.grid(row=6, column=0, columnspan=2, pady=10)

    # Tabla para mostrar resultados
    columnas = ("Sector", "Nombre del Sector", "Fecha de Visita", "Clasificación", "TIPO CLIENTE")
    tabla = ttk.Treeview(root, columns=columnas, show="headings")
    for col in columnas:
        tabla.heading(col, text=col)
    tabla.grid(row=7, column=0, columnspan=2, padx=10, pady=5)


     # Marco para contener la tabla y el scroll
    frame_tabla = ttk.Frame(root)
    frame_tabla.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

    # Scrollbars
    scrollbar_y = ttk.Scrollbar(frame_tabla, orient="vertical")
    scrollbar_x = ttk.Scrollbar(frame_tabla, orient="horizontal")

    columnas = ("Sector", "Nombre del Sector", "Fecha de Visita", "Clasificación", "TIPO CLIENTE")
    tabla = ttk.Treeview(frame_tabla, columns=columnas, show="headings", yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    for col in columnas:
        tabla.heading(col, text=col)
        tabla.column(col, width=200)  # Ajusta el ancho de las columnas

    tabla.pack(side="top", fill="both", expand=True)
    
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x.pack(side="bottom", fill="x")

    scrollbar_y.config(command=tabla.yview)
    scrollbar_x.config(command=tabla.xview)

    # Entrada para correo
    ttk.Label(root, text="Correo Destino:").grid(row=6, column=0, padx=10, pady=5)
    entrada_correo = ttk.Entry(root)
    entrada_correo.grid(row=9, column=0, padx=10, pady=5)

    btn_enviar = ttk.Button(root, text="📧 Generar Reporte", command=enviar_correo)
    btn_enviar.grid(row=9, column=1)









