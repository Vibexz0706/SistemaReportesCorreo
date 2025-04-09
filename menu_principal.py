import tkinter as tk
from tkinter import ttk
from reportes_individuales import abrir_sistema_reportes
from reportes_totales import abrir_reporte_por_pais
from reportes_pais import abrir_sistema_pais

def salir():
    ventana_inicio.quit()

# Crear ventana principal
ventana_inicio = tk.Tk()
ventana_inicio.title("📊 Sistema de Reportes")
ventana_inicio.geometry("500x300")

ttk.Label(ventana_inicio, text="Bienvenido al Sistema", font=("Arial", 16)).pack(pady=20)

# Botones para abrir reportes
btn_individual = ttk.Button(ventana_inicio, text="📋 Reportes Individuales", command=abrir_sistema_reportes )
btn_individual.pack(pady=10)

btn_pais = ttk.Button(ventana_inicio, text="🌎 Reportes por País", command=abrir_sistema_pais)
btn_pais.pack(pady=10)

btn_totales = ttk.Button(ventana_inicio, text="📊 Reportes Totales", command=abrir_reporte_por_pais)
btn_totales.pack(pady=10)

btn_salir = ttk.Button(ventana_inicio, text="❌ Salir", command=salir)
btn_salir.pack(pady=10)

ventana_inicio.mainloop()   
