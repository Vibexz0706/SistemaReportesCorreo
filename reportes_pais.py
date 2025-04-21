import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from fpdf import FPDF
import win32com.client as win32
from datetime import datetime, timedelta
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
        "50606": "JOHAN ALBERTO ARCE NÚÑEZ",
        "50704": "LOURDES ROA",
        "50706": "NAZARETH NAVARRO"
        # Agrega más según lo que necesites
    }
    df_extra["NOMBRE_SECTOR"] = df_extra["SECTOR"].map(mapeo_nombres)

    # Extraer número de horas del texto como "8 horas"
    df_extra["HORAS"] = df_extra["PERÍODO"].apply(lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)

    # Suma total de horas por nombre
    #resumen_horas = df_extra.groupby("NOMBRE_SECTOR")["HORAS"].sum().reset_index()
    # Filtrar solo horas >= 6
    df_filtrado = df_extra[df_extra["HORAS"] >= 6]

    # Agrupar y sumar
    resumen_horas = df_filtrado.groupby("NOMBRE_SECTOR")["HORAS"].sum().reset_index()



   # resumen_horas = df_extra[df_extra["HORAS"] >= 6].groupby("NOMBRE_SECTOR")["HORAS"].sum().reset_index()

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

    global df_filtrado
  
    """ 
    fechas_str = ["21/04/2025", "22/04/2025", "23/04/2025", "24/04/2025", "25/04/2025", "28/04/2025", "29/04/2025", "30/04/2025",
                  "01/05/2025","02/05/2025","05/05/2025","06/05/2025","07/05/2025","08/05/2025","09/05/2025","12/05/2025","13/05/2025","14/05/2025","15/04/2025","16/05/2025"]
    fechas = [datetime.strptime(f, "%d/%m/%Y") for f in fechas_str]
    #fecha_hoy = datetime.strptime("10/04/2025", "%d/%m/%Y")  # o usa datetime.today() si quieres la real
    fecha_hoy = datetime.today().strftime("%d/%m/%Y")

    # Contar cuántas fechas son <= fecha_hoy
    global conteo
    conteo = sum(1 for f in fechas if f <= fecha_hoy)
    """

    """

    CICLOS = {
        "CICLO 1": (datetime(2025, 1, 1), datetime(2025, 2, 10)),
        "CICLO 2": (datetime(2025, 2, 11), datetime(2025, 3, 13)),
        "CICLO 3": (datetime(2025, 3, 14), datetime(2025, 4, 10)),
        "CICLO 4": (datetime(2025, 4, 21), datetime(2025, 6, 16)),
    }

    def obtener_ciclo_actual(fecha_actual):
        for ciclo, (inicio, fin) in CICLOS.items():
            if inicio <= fecha_actual <= fin:
                return ciclo, inicio, fin
        return None, None, None

    fecha_actual = datetime.now()
    ciclo_actual, fecha_ciclo_inicio, fecha_ciclo_fin = obtener_ciclo_actual(fecha_actual)
    if not ciclo_actual:
        messagebox.showerror("Error", "No hay datos para el ciclo actual, actualiza la base.")
        return

    print(f"Ciclo actual: {ciclo_actual} ({fecha_ciclo_inicio.date()} a {fecha_ciclo_fin.date()})")


    try:
        df = pd.read_excel("base.xlsx")  # Ajusta el nombre del archivo
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True)

    #   Filtrar por fechas del ciclo actual
        df_filtrado = df[(df["FECHA"] >= fecha_ciclo_inicio) & (df["FECHA"] <= fecha_ciclo_fin)]

        if df_filtrado.empty:
            messagebox.showerror("Error", f"No hay datos disponibles para {ciclo_actual}.\n({fecha_ciclo_inicio.date()} a {fecha_ciclo_fin.date()})\nVerifica o actualiza la base de datos.")
            return

        print(f"✅ Datos encontrados para {ciclo_actual}: {len(df_filtrado)} registros")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar o procesar el archivo.\nDetalles: {e}")
        return
    """

    feriados = {
        datetime(2025, 3, 19),
        datetime(2025, 4, 11),
        datetime(2025, 4, 14),
        datetime(2025, 4, 15),
        datetime(2025, 4, 16),
        datetime(2025, 4, 17),
        datetime(2025, 4, 18),
        datetime(2025, 5, 1),
        # Agrega más feriados según lo necesites
    }

    global conteo
    conteo = 0

    def generar_ciclo(inicio, duracion=20):
        fechas = []
        actual = inicio
        while len(fechas) < duracion:
            if actual.weekday() < 5 and actual not in feriados:  # Día hábil
                fechas.append(actual)
            actual += timedelta(days=1)
        return fechas
    

    # Generar todos los ciclos del año automáticamente
    def generar_ciclos_del_anio(inicio_2025, cantidad_ciclos=20):
        ciclos = []
        inicio = inicio_2025
        for _ in range(cantidad_ciclos):
            ciclo = generar_ciclo(inicio)
            ciclos.append(ciclo)
            # El siguiente ciclo comienza el siguiente día hábil después del último del ciclo actual
            siguiente_inicio = ciclo[-1] + timedelta(days=1)
            while siguiente_inicio.weekday() >= 5 or siguiente_inicio in feriados:
                siguiente_inicio += timedelta(days=1)
            inicio = siguiente_inicio
        return ciclos
    
    # Crear los ciclos para 2025 (ajusta la fecha inicial si cambia)
    ciclos = generar_ciclos_del_anio(datetime(2025, 3, 14))

    # Fecha de hoy (puedes usar datetime.today() si quieres que sea automática)
    fecha_hoy = datetime.strptime("21/04/2025", "%d/%m/%Y")

    # Encontrar el ciclo actual y contar los días transcurridos
    for idx, ciclo in enumerate(ciclos):
        if ciclo[0] <= fecha_hoy <= ciclo[-1]:
            global DIA_DEL_CICLO
            DIA_DEL_CICLO = sum(1 for f in ciclo if f <= fecha_hoy)
            print(f"Estamos en el ciclo #{idx+1}, han transcurrido {DIA_DEL_CICLO} días hábiles.")
            break
    else:
        print("La fecha actual no está dentro de ningún ciclo definido.")


#////////////////////////////////////////////////////////////////////////////////////////////////////CAMBIAR CONTEO POR DIA_DEL_CICLO///////////////////////////////////////////
    
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
    
    

    df_extra = pd.read_excel(archivo_adicional, usecols=["SECTOR CLIENTE", "PERÍODO", "FECHA_INCLUSION"],dtype={"SECTOR CLIENTE": str})
    df_extra["SECTOR CLIENTE"] = (
        df_extra["SECTOR CLIENTE"]
        .astype(str)
        .str.encode("ascii", "ignore")   # elimina caracteres raros
        .str.decode("utf-8")
        .str.strip()
    )
# Limpiar columnas justo después de cargar el Excel
    df_extra.columns = df_extra.columns.str.strip()

# Forzar nombre correcto si viene como "SECTOR CLIENTE" con espacios
    for col in df_extra.columns:
        if "SECTOR" in col.upper() and "CLIENTE" in col.upper():
            df_extra.rename(columns={col: "SECTOR CLIENTE"}, inplace=True)

    print("✅ Filas leídas de df_extra:", len(df_extra))
    print("📌 Primeras filas:")
    print(df_extra.head())
    print("📌 Columnas:", df_extra.columns.tolist())
    print("🧪 Tipos originales:")
    print(df_extra.dtypes)
    print("📌 Ejemplo de valor:", df_extra["SECTOR CLIENTE"].iloc[0], type(df_extra["SECTOR CLIENTE"].iloc[0]))
    for val in df_extra["SECTOR CLIENTE"].unique():
        print(f"➡️ '{val}' (len={len(val)})")



    # Normalización correcta
    df_extra.columns = df_extra.columns.str.strip()
    df_extra["SECTOR CLIENTE"] = df_extra["SECTOR CLIENTE"].astype(str).str.strip()
    df_extra["PERÍODO"] = df_extra["PERÍODO"].astype(str).str.strip()
    df_extra["FECHA_INCLUSION"] = pd.to_datetime(df_extra["FECHA_INCLUSION"], errors="coerce", dayfirst=True)

    """
    # Filtrar solo registros dentro del ciclo actual
    df_extra = df_extra[
        (df_extra["FECHA_INCLUSION"] >= fecha_ciclo_inicio) &
        (df_extra["FECHA_INCLUSION"] <= fecha_ciclo_fin)
    ]
    """
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

        total_horas = df_extra_filtrado[df_extra_filtrado["HORAS"] >= 6]["HORAS"].sum()

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
    pdf.cell(105, 10, "Cobertura Real", 1, 0, "C")
    pdf.cell(100, 10, "Datos Reales", 1, 0, "C")
    pdf.cell(75, 10, "Cobertura Ciclo", 1, 1, "C")
    pdf.cell(55, 10, "", 1, 0, "C")
    pdf.cell(25, 10, "MÉDICOS", 1, 0, "C")
    pdf.cell(25, 10, "FARMACIAS", 1, 0, "C")
    pdf.cell(35, 10, "MÉDICOS", 1, 0, "C")
    pdf.cell(35, 10, "FARMACIAS", 1, 0, "C")
    pdf.cell(35, 10, "VIP", 1, 0, "C")
    pdf.cell(25, 10, "MÉDICOS", 1, 0, "C")
    pdf.cell(25, 10, "ESTANDARES", 1, 0, "C")
    pdf.cell(25, 10, "FARMACIAS", 1, 0, "C")
    pdf.cell(25, 10, "VIP", 1, 0, "C")
    pdf.cell(25, 10, "MÉDICOS", 1, 0, "C")
    pdf.cell(25, 10, "VIP", 1, 0, "C")
    pdf.cell(25, 10, "ESTANDARES", 1, 1, "C")

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
        total_medicos_estandar = df_consultor[df_consultor["CLASIFICACIÓN"].isin(["ESTANDAR"])].shape[0]  #///////////////////////////
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
            var_prueba = 7.5 * (total_dias - restar_horas / 8)
            var_cobertura_real = 7.5 * conteo
            var_cobertura_real_vip = 1.8 * conteo
            var_cobertura_estandar = 5.7 * conteo
            
        elif consultor == "RICARDO  ZUNIGA":
            var_arreglo_real = 9 * (total_dias - restar_horas / 8)
            var_arreglo_regentes = 4 * (total_dias - restar_horas / 8)
            var_arreglo_vip = 2.8 * (total_dias - restar_horas / 8)
            var_prueba = (total_dias - restar_horas / 8)
            var_cobertura_real = 9 * conteo
            var_cobertura_real_vip = 2.8 * conteo
            var_cobertura_estandar = 6.2 * conteo
        elif consultor == "KAROLINA MONTERO":
            var_arreglo_real = 9 * (total_dias - restar_horas / 8)
            var_arreglo_regentes = 4 * (total_dias - restar_horas / 8)
            var_arreglo_vip = 2.2 * (total_dias - restar_horas / 8)
            var_prueba = (total_dias - restar_horas / 8)
            var_cobertura_real = 9 * conteo
            var_cobertura_real_vip = 2.2 * conteo
            var_cobertura_estandar = 6.8 * conteo
        elif consultor == "KEREN CARVAJAL":
            var_arreglo_real = 9 * (total_dias - restar_horas / 8)
            var_arreglo_regentes = 4 * (total_dias - restar_horas / 8)
            var_arreglo_vip = 2.05 * (total_dias - restar_horas / 8)
            var_prueba = (total_dias - restar_horas / 8)
            var_cobertura_real = 9 * conteo
            var_cobertura_real_vip = 2.05 * conteo     
            var_cobertura_estandar = 7.05 * conteo

        elif consultor == "JOHAN ALBERTO ARCE NÚÑEZ":
            var_arreglo_real = 9 * (total_dias - restar_horas / 8)
            var_arreglo_regentes = 4 * (total_dias - restar_horas / 8)
            var_arreglo_vip = 2.35 * (total_dias - restar_horas / 8)
            var_prueba = (total_dias - restar_horas / 8)
            var_cobertura_real = 9 * conteo
            var_cobertura_real_vip = 2.35 * conteo
            var_cobertura_estandar = 6.6 * conteo
        elif consultor == "LOURDES ROA":
            var_arreglo_real = 9 * (total_dias - restar_horas / 8)
            var_arreglo_regentes = 4 * (total_dias - restar_horas / 8)
            var_arreglo_vip = 2.35 * (total_dias - restar_horas / 8)
            var_prueba = (total_dias - restar_horas / 8)
            var_cobertura_real = 9 * conteo
            var_cobertura_real_vip = 2.35 * conteo
            var_cobertura_estandar = 6.6 * conteo
        elif consultor == "NAZARETH NAVARRO":
            var_arreglo_real = 9 * (total_dias - restar_horas / 8)
            var_arreglo_regentes = 4 * (total_dias - restar_horas / 8)
            var_arreglo_vip = 2.35 * (total_dias - restar_horas / 8)
            var_prueba = (total_dias - restar_horas / 8)
            var_cobertura_real = 9 * conteo
            var_cobertura_real_vip = 2.35 * conteo
            var_cobertura_estandar = 6.6 * conteo
        else:
            raise ValueError(f"Consultor no reconocido: {consultor}")

        cobertura_medicos_pct = round((var_arreglo / var_arreglo_real) * 100, 2) if var_arreglo_real > 0 else 0
        cobertura_regentes_pct = round((total_regentes / var_arreglo_regentes) * 100, 2) if var_arreglo_regentes > 0 else 0
        cobertura_vip_pct = round((total_vip / var_arreglo_vip) * 100, 2) if var_arreglo_real > 0 else 0
        # 🔍 Mostrar en consola detalles del cálculo
        print(f"\n--- Cálculos para {consultor} ---")
        print(f"Total médicos: {total_medicos}")
        print(f"Total regentes: {total_regentes}")
        print(f"Total VIP: {total_vip}")
        print(f"Total días reales: {total_dias}")
        
        print(f"RESTAR_HORAS: {restar_horas}")
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
        if cobertura_medicos_pct < 95:
            pdf.set_text_color(255, 0, 0)  # Rojo
            pdf.cell(35, 10, f"[RED]{cobertura_medicos_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro
        elif cobertura_medicos_pct > 94.99 and cobertura_medicos_pct < 99:
            pdf.set_text_color(255, 255, 0)  # Amarillo
            pdf.cell(35, 10, f"[AMARILLO]{cobertura_medicos_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro
        else:
            pdf.set_text_color(0, 255, 0)  # verde
            pdf.cell(35, 10, f"[VERDE]{cobertura_medicos_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro
        
        if cobertura_regentes_pct < 95:
            pdf.set_text_color(255, 0, 0)  # Rojo
            pdf.cell(35, 10, f"[RED]{cobertura_regentes_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro
        elif cobertura_regentes_pct > 94.99 and cobertura_regentes_pct < 99:
            pdf.set_text_color(255, 255, 0)  # Amarillo
            pdf.cell(35, 10, f"[AMARILLO]{cobertura_regentes_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro
        else:
            pdf.set_text_color(0, 255, 0)  # verde
            pdf.cell(35, 10, f"[VERDE]{cobertura_regentes_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro

        if cobertura_vip_pct < 95:
            pdf.set_text_color(255, 0, 0)  # Rojo
            pdf.cell(35, 10, f"[RED]{cobertura_vip_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro
        elif cobertura_vip_pct > 95 and cobertura_vip_pct < 99:
            pdf.set_text_color(255, 255, 0)  # Amarillo
            pdf.cell(35, 10, f"[AMARILLO]{cobertura_vip_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro
        else: 
            pdf.set_text_color(0, 255, 0)  # verde
            pdf.cell(35, 10, f"[VERDE]{cobertura_vip_pct:.2f}%", 1, 0, "C")
            pdf.set_text_color(0, 0, 0)    # Restaurar a negro
        pdf.cell(25, 10, str(total_medicos_vip_estandar), 1, 0, "C")
        pdf.cell(25, 10, str(total_medicos_estandar), 1, 0, "C")
        pdf.cell(25, 10, str(total_medicos_regentes), 1, 0, "C")
        pdf.cell(25, 10, str(total_medicos_vip), 1, 0, "C")
        pdf.cell(25, 10, str(var_cobertura_real), 1, 0, "C")
        pdf.cell(25, 10, f"{var_cobertura_real_vip:.2f}", 1, 0, "C")
        pdf.cell(25, 10, f"{var_cobertura_estandar:.2f}", 1, 1, "C")

    # Guardar el PDF
    pdf.output(ruta_pdf)
    messagebox.showinfo("PDF Generado", f"El reporte se ha guardado en:\n{ruta_pdf}")
    return ruta_pdf

def extraer_todas_las_tablas_pdf(ruta_pdf):
    import pdfplumber

    tabla_html = ""

    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            tablas = page.extract_tables()

            if tablas:
                for i, tabla in enumerate(tablas):
                    #tabla_html += f"<h3>Resumen País {i + 1}</h3>"
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

                            estilo = "padding: 5px; border: 1px solid #000;"  # estilo base
                            if "[RED]" in celda:
                                estilo += "color: red; font-weight: bold;"
                                celda = celda.replace("[RED]", "")
                            elif "[AMARILLO]" in celda:
                                estilo += "color: orange; font-weight: bold;"
                                celda = celda.replace("[AMARILLO]", "")
                            elif "[VERDE]" in celda:
                                estilo += "color: green; font-weight: bold;"
                                celda = celda.replace("[VERDE]", "")

                            if fila_index == 0:
                                # Manejo de encabezados con colspan
                                colspan = 1
                                for next_col in fila[col_index + 1:]:
                                    if next_col is None:
                                        colspan += 1
                                    else:
                                        break
                                if colspan > 1:
                                    tabla_html += f"<th colspan='{colspan}' style='{estilo}'>{celda}</th>"
                                    skip_celdas = colspan - 1
                                else:
                                    tabla_html += f"<th style='{estilo}'>{celda}</th>"
                            else:
                                tabla_html += f"<td style='{estilo}'>{celda}</td>"

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
    <p>Estimada Eyilde, Buenos Días
    <p>A continuación le indicamos la situación de su equipo, según la información cargada al sistema.</p>
    <p>Se procesó información de los consultores.</p>
    <p>Día Ciclo:{conteo}</p> 
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


