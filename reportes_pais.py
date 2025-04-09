import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from fpdf import FPDF
import win32com.client as win32
from datetime import datetime
import numpy as np
import pdfplumber
import re
import unicodedata




# ======================== CONFIGURACIÓN ========================
ruta_carpeta = r"C:\Users\jabar\OneDrive\Escritorio\PruebaFiltro"
df_filtrado = None  # Variable global para almacenar los datos filtrados

# ======================== FUNCIONES ========================
def obtener_ultimo_excel():
    nombre_archivo = "base.xlsx"  # <- Cambialo por el nombre real del archivo
    ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)
    return ruta_archivo if os.path.exists(ruta_archivo) else None
    #archivos = glob.glob(os.path.join(ruta_carpeta, "*.xlsx"))
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
        "50603": "RICARDO  ZUNIGA",
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

def obtener_excel_especifico(nombre_archivo):
    archivo_path = os.path.join(ruta_carpeta, nombre_archivo)
    return archivo_path if os.path.exists(archivo_path) else None

def aplicar_filtros():
    CICLOS = {
        "CICLO 1": (datetime(2025, 1, 1), datetime(2025, 2, 10)),
        "CICLO 2": (datetime(2025, 2, 11), datetime(2025, 3, 13)),
        "CICLO 3": (datetime(2025, 3, 14), datetime(2025, 4, 12)),
        "CICLO 4": (datetime(2025, 4, 13), datetime(2025, 5, 15)),
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

    fecha_inicio = entrada_fecha_inicio.get().strip()
    fecha_fin = entrada_fecha_fin.get().strip()
    pais_seleccionado = combo_pais.get()

    if fecha_inicio:
        df = df[df["FECHA DE VISITA"] >= pd.to_datetime(fecha_inicio, errors="coerce", dayfirst=True)]
    if fecha_fin:
        df = df[df["FECHA DE VISITA"] <= pd.to_datetime(fecha_fin, errors="coerce", dayfirst=True)]

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

    df_filtrado = df.copy()
    for row in tabla.get_children():
        tabla.delete(row)

    if df.empty:
        messagebox.showwarning("Sin Resultados", "No se encontraron registros con esos filtros.")
    else:
        for _, fila in df.iterrows():
            tabla.insert("", "end", values=(fila["SECTOR"], fila["NOMBRE DEL SECTOR"], fila["FECHA DE VISITA"].strftime('%d/%m/%Y'), fila["CLASIFICACIÓN"], fila["TIPO CLIENTE"]))
        total_registros = len(df)
        tabla.insert("", "end", values=("TOTAL", "", "", total_registros, ""))

    # === CALCULAR HORAS A RESTAR DEL SEGUNDO EXCEL ===
    archivo_adicional = obtener_excel_especifico("ExportacaoAE_DNT-Ciclo 3.xlsx")
    if not archivo_adicional:
        return

    df_extra = pd.read_excel(archivo_adicional, usecols=["SECTOR CLIENTE", "PERÍODO", "FECHA_INCLUSION"])
    print("✅ Filas leídas de df_extra:", len(df_extra))
    print("📌 Primeras filas:")
    print(df_extra.head())
    print("📌 Columnas:", df_extra.columns.tolist())

    # Normalización correcta
    df_extra.columns = df_extra.columns.str.strip()
    df_extra["SECTOR CLIENTE"] = df_extra["SECTOR CLIENTE"].astype(str).str.strip()
    df_extra["PERÍODO"] = df_extra["PERÍODO"].astype(str).str.strip()
    df_extra["FECHA_INCLUSION"] = pd.to_datetime(df_extra["FECHA_INCLUSION"], errors="coerce", dayfirst=True)

    # Filtrar solo registros dentro del ciclo actual
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

    def normalizar(texto):
        return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8').strip().upper().replace("  ", " ")

    def extraer_horas(periodo):
        if pd.isnull(periodo):
            return 0
        texto = str(periodo).strip().upper()
        match = re.search(r'(\d+)', texto)
        if match:
            return int(match.group(1))
        else:
            print(f"[WARNING] No se encontró número en PERÍODO: '{texto}'")
            return 0

    global HORAS_POR_CONSULTOR
    HORAS_POR_CONSULTOR = {}

    nombres_en_df = df_filtrado["NOMBRE DEL SECTOR"].dropna().unique()
    for nombre in nombres_en_df:
        nombre_normalizado = normalizar(nombre)
        codigo_sector = None

        for codigo, nombre_mapeado in mapeo_nombres.items():
            if nombre_normalizado == normalizar(nombre_mapeado):
                codigo_sector = codigo
                break

        if not codigo_sector:
            print(f"[ADVERTENCIA] No se encontró código para el consultor '{nombre}'")
            continue

        # Verificación extra
        print(f"\n🔍 DEBUG - Consultor: {nombre_normalizado}")
        print(f"Buscando código: {codigo_sector}")
        print("👉 Valores únicos actuales en SECTOR CLIENTE:", df_extra["SECTOR CLIENTE"].unique())

        df_extra_filtrado = df_extra[df_extra["SECTOR CLIENTE"] == codigo_sector].copy()

        if df_extra_filtrado.empty:
            print(f"[INFO] No hay filas en df_extra para el código {codigo_sector}")
            continue

        df_extra_filtrado["HORAS"] = df_extra_filtrado["PERÍODO"].apply(extraer_horas)

        total_horas = df_extra_filtrado["HORAS"].sum()
        HORAS_POR_CONSULTOR[nombre_normalizado] = total_horas

        print(f"\n✅ Detalle para {nombre_normalizado} (código {codigo_sector}):")
        print(df_extra_filtrado[["FECHA_INCLUSION", "PERÍODO", "HORAS"]])

    print("\n=== HORAS A RESTAR POR CONSULTOR ===")
    for nombre, horas in HORAS_POR_CONSULTOR.items():
        print(f"{nombre}: {horas} horas")




#////////////////////////////////////////////////////////////////////////ESTE ES EL REPORTE DE ESTA SECCION///////////////////
def generar_pdf_resumen():
    if df_filtrado is None or df_filtrado.empty:
        messagebox.showwarning("Advertencia", "No hay datos filtrados para generar el PDF.")
        return

    ruta_pdf = os.path.join(ruta_carpeta, "reporte_resumen.pdf")
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font("Arial", "B", 14)
    pdf.cell(300, 10, "Resumen de Consultores", ln=True, align="C")
    pdf.ln(10)

    # Fecha de rango global del filtro
    df_filtrado["FECHA DE VISITA"] = pd.to_datetime(df_filtrado["FECHA DE VISITA"])
    fecha_inicio = df_filtrado["FECHA DE VISITA"].min().date()
    fecha_fin = df_filtrado["FECHA DE VISITA"].max().date()

    pdf.set_font("Arial", "B", 12)
    pdf.cell(300, 10, f"Datos desde {fecha_inicio.strftime('%d-%m-%Y')} hasta {fecha_fin.strftime('%d-%m-%Y')}", ln=True, align="C")
    pdf.ln(10)

    # Encabezados
    pdf.set_font("Arial", "B", 10)
    pdf.cell(55, 10, "Consultores", 1, 0, "C")
    pdf.cell(50, 10, "Promedio Visitas", 1, 0, "C")
    pdf.cell(75, 10, "Cobertura Real", 1, 0, "C")
    pdf.cell(50, 10, "Programados", 1, 0, "C")
    pdf.cell(25, 10, "Panel", 1, 1, "C")
    pdf.cell(55, 10, "", 1, 0, "C")
    pdf.cell(25, 10, "MÉDICOS", 1, 0, "C")
    pdf.cell(25, 10, "FARMACIAS", 1, 0, "C")
    pdf.cell(25, 10, "Médicos", 1, 0, "C")
    pdf.cell(25, 10, "Farmacias", 1, 0, "C")
    pdf.cell(25, 10, "Médicos VIP", 1, 0, "C")
    pdf.cell(25, 10, "Médicos", 1, 0, "C")
    pdf.cell(25, 10, "Farmacias", 1, 0, "C")
    pdf.cell(25, 10, "Médicos VIP", 1, 1, "C")

    pdf.set_font("Arial", "", 10)

    consultores_unicos = df_filtrado["NOMBRE DEL SECTOR"].dropna().unique()

    for consultor in consultores_unicos:
        df_consultor = df_filtrado[df_filtrado["NOMBRE DEL SECTOR"] == consultor]

        total_medicos = len(df_consultor)
        total_vip = len(df_consultor[df_consultor["CLASIFICACIÓN"] == "VIP"])
        total_regentes = len(df_consultor[df_consultor["CLASIFICACIÓN"].isna()])
        var_arreglo = total_medicos - total_regentes 

        total_medicos_vip_estandar = df_consultor[df_consultor["CLASIFICACIÓN"].isin(["VIP", "ESTANDAR"])].shape[0]
        total_medicos_vip = df_consultor[df_consultor["CLASIFICACIÓN"].isin(["VIP"])].shape[0]
        total_medicos_regentes = df_consultor[df_consultor["TIPO CLIENTE"].isin(["PDV"])].shape[0]

        df_consultor["FECHA DE VISITA"] = pd.to_datetime(df_consultor["FECHA DE VISITA"])
        fechas_unicas = sorted(df_consultor["FECHA DE VISITA"].dropna().unique())
        total_dias = len(fechas_unicas) if len(fechas_unicas) > 0 else 1  

        cobertura_medicos = round(var_arreglo / total_dias, 2)
        cobertura_regentes = round(total_regentes / total_dias, 2)

        nombre_normalizado = unicodedata.normalize('NFKD', consultor).encode('ASCII', 'ignore').decode('utf-8').strip().upper().replace("  ", " ")
        restar_horas = HORAS_POR_CONSULTOR.get(nombre_normalizado, 0)

        # 🔁 Calcular cobertura real (porcentajes), dependiendo del consultor
        if consultor == "ANA KAREN MORA":
            var_arreglo_real = 7.5 * (total_dias - restar_horas / 8)
            var_arreglo_regentes = 4 * (total_dias - restar_horas / 8)
            var_arreglo_vip = 1.8 * (total_dias - restar_horas / 8)
        else:
            var_arreglo_real = 9 * (total_dias - restar_horas / 8)
            var_arreglo_regentes = 6 * (total_dias - restar_horas / 8)
            var_arreglo_vip = 4 * (total_dias - restar_horas / 8)

        cobertura_medicos_pct = round((var_arreglo / var_arreglo_real) * 100, 2) if var_arreglo_real > 0 else 0
        cobertura_regentes_pct = round((total_regentes / var_arreglo_regentes) * 100, 2) if var_arreglo_regentes > 0 else 0
        cobertura_vip_pct = round((total_vip / var_arreglo_vip) * 100, 2) if var_arreglo_real > 0 else 0
        # 🔍 Mostrar en consola detalles del cálculo
        print(f"\n--- Cálculos para {consultor} ---")
        print(f"Total médicos: {total_medicos}")
        print(f"Total regentes: {total_regentes}")
        print(f"Total VIP: {total_vip}")
        print(f"Total días reales: {total_dias}")
        #print(f"RESTAR_HORAS: {RESTAR_HORAS}")
        print(f"var_arreglo (Médicos - Regentes): {var_arreglo}")
        print(f"var_arreglo_real (objetivo médicos): {var_arreglo_real}")
        print(f"var_arreglo_regentes (objetivo regentes): {var_arreglo_regentes}")
        print(f"Cobertura Médicos diaria: {cobertura_medicos}")
        print(f"Cobertura Regentes diaria: {cobertura_regentes}")
        print(f"% Cobertura Médicos: {cobertura_medicos_pct}%")
        print(f"% Cobertura Regentes: {cobertura_regentes_pct}%")
        print(f"% Cobertura VIP: {cobertura_vip_pct}%")
        print(f"Programados médicos+vip: {total_medicos_vip_estandar}")
        print(f"Programados regentes: {total_medicos_regentes}")
        print(f"VIP reales: {total_medicos_vip}")

        # 🧾 Fila en el PDF
        pdf.cell(55, 10, consultor, 1, 0, "C")
        pdf.cell(25, 10, str(cobertura_medicos), 1, 0, "C")
        pdf.cell(25, 10, str(cobertura_regentes), 1, 0, "C")
        pdf.cell(25, 10, f"{cobertura_medicos_pct:.2f}%", 1, 0, "C")
        pdf.cell(25, 10, f"{cobertura_regentes_pct:.2f}%", 1, 0, "C")
        pdf.cell(25, 10, f"{cobertura_vip_pct:.2f}%", 1, 0, "C")
        pdf.cell(25, 10, str(total_medicos_vip_estandar), 1, 0, "C")
        pdf.cell(25, 10, str(total_medicos_regentes), 1, 0, "C")
        pdf.cell(25, 10, str(total_medicos_vip), 1, 1, "C")
    # Guardar el PDF
    pdf.output(ruta_pdf)
    messagebox.showinfo("PDF Generado", f"El reporte se ha guardado en:\n{ruta_pdf}")
    return ruta_pdf


#-------------------------------------------------------------------------
def generar_pdf_por_pais():
    if df_filtrado is None or df_filtrado.empty:
        messagebox.showwarning("Advertencia", "No hay datos filtrados para generar el PDF.")
        return

    ruta_pdf = os.path.join(ruta_carpeta, "reporte_resumen_por_pais.pdf")
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font("Arial", "B", 14)
    pdf.cell(300, 10, "Resumen de Consultores por País", ln=True, align="C")
    pdf.ln(10)

    # Asegurar que SECTOR sea string y limpiar espacios
    df_filtrado["SECTOR"] = df_filtrado["SECTOR"].astype(str).str.strip()

    # Verificar valores únicos en SECTOR antes de filtrar
    print("📊 Valores únicos en 'SECTOR':", df_filtrado["SECTOR"].unique())

    # Definir los códigos de país
    paises = {"506": "Costa Rica", "507": "Panamá", "505": "Nicaragua"}

    for codigo_pais, nombre_pais in paises.items():
        df_pais = df_filtrado[df_filtrado["SECTOR"].str[:3] == codigo_pais]  # Filtrar por los 3 primeros dígitos

        if df_pais.empty:
            print(f"⚠️ No hay datos para {nombre_pais} ({codigo_pais})")
            continue

        # Agregar título del país
        pdf.set_font("Arial", "B", 12)
        pdf.cell(300, 10, f"{nombre_pais}", ln=True, align="C")
        pdf.ln(5)

        # Encabezado de la tabla
        pdf.set_font("Arial", "B", 10)
        pdf.cell(55, 10, "CONSULTORES", 1, 0, "C")
        pdf.cell(80, 10, "PROMEDIO VISITAS", 1, 0, "C")
        pdf.cell(120, 10, "Programados", 1, 1, "C")

        pdf.cell(55, 10, "", 1, 0, "C")
        pdf.cell(40, 10, "MÉDICOS", 1, 0, "C")
        pdf.cell(40, 10, "FARMACIAS", 1, 0, "C")
        pdf.cell(40, 10, "MÉDICOS", 1, 0, "C")
        pdf.cell(40, 10, "Farmacias", 1, 0, "C")
        pdf.cell(40, 10, "Médio VIP", 1, 1, "C")

        pdf.set_font("Arial", "", 10)

        consultores_unicos = df_pais["NOMBRE DEL SECTOR"].dropna().unique()

        for consultor in consultores_unicos:
            df_consultor = df_pais[df_pais["NOMBRE DEL SECTOR"] == consultor]

            total_medicos = len(df_consultor)
            total_regentes = len(df_consultor[df_consultor["CLASIFICACIÓN"].isna()])
            var_arreglo = total_medicos - total_regentes

            total_medicos_vip_estandar = df_consultor[df_consultor["CLASIFICACIÓN"].isin(["VIP", "ESTANDAR"])].shape[0]
            total_medicos_vip = df_consultor[df_consultor["CLASIFICACIÓN"].isin(["VIP"])].shape[0]
            total_medicos_regentes = df_consultor[df_consultor["TIPO CLIENTE"].isin(["PDV"])].shape[0]

            df_consultor["FECHA DE VISITA"] = pd.to_datetime(df_consultor["FECHA DE VISITA"])
            fechas_unicas = sorted(df_consultor["FECHA DE VISITA"].dropna().unique())
            total_dias = len(fechas_unicas) if len(fechas_unicas) > 0 else 1  

            cobertura_medicos = round(var_arreglo / total_dias, 2)
            cobertura_regentes = round(total_regentes / total_dias, 2)

            # Agregar fila a la tabla
            pdf.cell(55, 10, consultor, 1, 0, "C")
            pdf.cell(40, 10, str(cobertura_medicos), 1, 0, "C")
            pdf.cell(40, 10, str(cobertura_regentes), 1, 0, "C")
            pdf.cell(40, 10, str(total_medicos_vip_estandar), 1, 0, "C")
            pdf.cell(40, 10, str(total_medicos_regentes), 1, 0, "C")
            pdf.cell(40, 10, str(total_medicos_vip), 1, 1, "C")

        pdf.ln(10)  # Espacio entre países

    pdf.output(ruta_pdf)
    messagebox.showinfo("PDF Generado", f"El reporte se ha guardado en:\n{ruta_pdf}")
    return ruta_pdf


#--------------------------------------------------------------------

def generar_pdf_sin_duplicados():
    if df_filtrado is None or df_filtrado.empty:
        messagebox.showwarning("Advertencia", "No hay datos filtrados para generar el PDF sin duplicados.")
        return

    ruta_pdf = os.path.join(ruta_carpeta, "reporte_sin_duplicados.pdf")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Título
    pdf.set_font("Arial", "B", 14)
    pdf.cell(190, 10, "Reporte de Consultores (Sin Duplicados)", ln=True, align="C")
    pdf.ln(10)

    # Obtener los nombres únicos de los consultores
    df_sin_duplicados = df_filtrado.drop_duplicates(subset=["NOMBRE DEL SECTOR"])

    # Crear tabla en PDF
    pdf.set_font("Arial", "B", 12)
    pdf.cell(100, 10, "Nombre del Consultor", 1, 1, "C")

    pdf.set_font("Arial", "", 10)
    for _, row in df_sin_duplicados.iterrows():
        pdf.cell(100, 10, row["NOMBRE DEL SECTOR"], 1, 1, "C")

    # Guardar el PDF
    pdf.output(ruta_pdf)
    messagebox.showinfo("PDF Generado", f"El reporte sin duplicados se ha guardado en:\n{ruta_pdf}")
    return ruta_pdf


    if df_filtrado is None or df_filtrado.empty:
        messagebox.showwarning("Advertencia", "No hay datos filtrados para generar el PDF.")
        return

    ruta_pdf = os.path.join(ruta_carpeta, "reporte_filtrado.pdf")
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    
    pdf.set_font("Arial", "B", 14)
    pdf.cell(270, 10, "Reporte Filtrado (Sin Duplicados)", ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("Arial", "B", 12)
    pdf.cell(60, 10, "Nombre del Sector", 1)
    pdf.cell(40, 10, "Fecha de Visita", 1)
    pdf.cell(40, 10, "Clasificación", 1)
    pdf.cell(40, 10, "Tipo Cliente", 1)
    pdf.ln()

    pdf.set_font("Arial", "", 10)
    for _, row in df_filtrado.iterrows():
        pdf.cell(60, 10, row["NOMBRE DEL SECTOR"], 1)
        pdf.cell(40, 10, row["FECHA DE VISITA"].strftime('%d/%m/%Y'), 1)
        pdf.cell(40, 10, row["CLASIFICACIÓN"], 1)
        pdf.cell(40, 10, row["TIPO CLIENTE"], 1)
        pdf.ln()

    pdf.output(ruta_pdf)
    messagebox.showinfo("PDF Generado", f"El reporte se ha guardado en:\n{ruta_pdf}")
    return ruta_pdf
#///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
import pdfplumber

import pdfplumber

def extraer_todas_las_tablas_pdf(ruta_pdf):
    import pdfplumber

    tabla_html = ""

    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            tablas = page.extract_tables()

            if tablas:
                for i, tabla in enumerate(tablas):
                    tabla_html += f"<h3>Resumen País {i + 1}</h3>"
                    tabla_html += "<table border='1' style='border-collapse: collapse; width: 100%; text-align: center;'>"

                    for fila_index, fila in enumerate(tabla):
                        tabla_html += "<tr>"
                        skip_celdas = 0

                        for col_index, celda in enumerate(fila):
                            if skip_celdas > 0:
                                skip_celdas -= 1
                                continue

                            if celda is None:
                                celda = ""

                            if fila_index == 0:  # Encabezado superior
                                # Detectar colspan hacia la derecha si hay None
                                colspan = 1
                                for next_col in fila[col_index + 1:]:
                                    if next_col is None:
                                        colspan += 1
                                    else:
                                        break

                                if colspan > 1:
                                    tabla_html += f"<th colspan='{colspan}'>{celda}</th>"
                                    skip_celdas = colspan - 1
                                else:
                                    tabla_html += f"<th>{celda}</th>"

                            else:  # Fila normal (datos o encabezado inferior)
                                tabla_html += f"<td>{celda}</td>"

                        tabla_html += "</tr>"

                    tabla_html += "</table><br>"

    return tabla_html if tabla_html else "<p>No se encontraron tablas en el PDF.</p>"



def es_correo_valido(correo):
    """Verifica si el correo tiene un formato válido usando una expresión regular."""
    patron = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(patron, correo) is not None
def enviar_correo():
    #ruta_pdf = generar_pdf_extra()
    correo_destino = entrada_correo.get().strip()
    if not correo_destino:
        messagebox.showwarning("Advertencia", "No se ingresó un correo. El reporte no se envió.")
        return
    if not es_correo_valido(correo_destino):
        messagebox.showerror("Error", "El correo ingresado no es válido. Inténtalo de nuevo.")
        return
    
    ruta_pdf = generar_pdf_resumen()

    tabla_html = extraer_todas_las_tablas_pdf(ruta_pdf)  # Extrae todas las tablas como HTML


    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    lista_destinatarios = ["xotoyip215@buides.com", "jose03pif@gmail.com"]
    mail.CC = "; ".join(lista_destinatarios)
    mail.To = correo_destino
    mail.Subject = "AGD - Resumen de su equipo"
    mail.HTMLBody = f"""
    <p>Estimada Francys, buenos días
    <p>A continuación le indicamos la situación de su equipo, según la información cargada al sistema.</p>
    <p>Se procesó información de los consultores.</p>
    <p>Día objetivo:</p> 
    {tabla_html}
    """
   # mail.Body = "Adjunto el reporte filtrado."
    mail.Attachments.Add(ruta_pdf)
    mail.Send()
    messagebox.showinfo("Correo Enviado", "El correo ha sido enviado correctamente.")

# ======================== VENTANAS ========================
def abrir_sistema_pais():
    #ventana_inicio.withdraw()
    global root, tabla, entrada_nombre, entrada_fecha_inicio, entrada_fecha_fin, combo_pais, combo_tipo_cliente, clasificacion_listbox, entrada_correo

    root = tk.Toplevel()
    root.title("Sistema de Reportes")
    root.geometry("1000x600")

    # Etiquetas y Entradas
    """
    ttk.Label(root, text="Nombre del Sector:").grid(row=0, column=0, padx=10, pady=5)
    entrada_nombre = ttk.Entry(root)
    entrada_nombre.grid(row=0, column=1, padx=10, pady=5)
"""
    ttk.Label(root, text="Fecha Inicio (dd/mm/yyyy):").grid(row=1, column=0, padx=10, pady=5)
    entrada_fecha_inicio = ttk.Entry(root)
    entrada_fecha_inicio.grid(row=1, column=1, padx=10, pady=5)

    ttk.Label(root, text="Fecha Fin (dd/mm/yyyy):").grid(row=2, column=0, padx=10, pady=5)
    entrada_fecha_fin = ttk.Entry(root)
    entrada_fecha_fin.grid(row=2, column=1, padx=10, pady=5)

    # Filtro de país (nuevo)
    ttk.Label(root, text="Seleccionar País:").grid(row=3, column=0, padx=10, pady=5)
    combo_pais = ttk.Combobox(root, values=["Todos", "Costa Rica", "Panamá", "Nicaragua", "Salvador", "Honduras", "Guatemala"])
    combo_pais.grid(row=3, column=1, padx=10, pady=5)
    combo_pais.current(0)  # Seleccionar "Todos" por defecto



    # Botón para aplicar filtros
    btn_filtrar = ttk.Button(root, text="🔍 Aplicar Filtros", command=aplicar_filtros)
    btn_filtrar.grid(row=6, column=0, columnspan=2, pady=10)

       # Botón para aplicar filtros
    #btn_filtrar = ttk.Button(root, text="🔍 Aplicar Filtros 2", command=aplicar_filtros_extra)   #--------------------------------
    #btn_filtrar.grid(row=6, column=0, columnspan=2, pady=10)

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
    entrada_correo.grid(row=10, column=0, padx=10, pady=5)

    if not entrada_correo:
        messagebox.showerror("ERROR", "Porfavor Ingresar un Correo")
    else:


        btn_enviar = ttk.Button(root, text="📧 Generar Reporte", command=enviar_correo)
        btn_enviar.grid(row=9, column=0)


 


   # btn_regresar = ttk.Button(root, text="⬅ Volver", command=lambda: (root.destroy(), ventana_inicio.deiconify()))
    #btn_regresar.grid(row=15, column=0)



# ======================== PÁGINA DE INICIO ========================

