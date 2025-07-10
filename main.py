import pandas as pd
import re

# Leer el archivo línea por línea
with open('rptRep_ReservacionesCompleto 14 a 20 de julio.csv', 'r', encoding='latin1') as file:
    lines = file.readlines()

# Lista para almacenar los registros procesados
datos_procesados = []

# Variables temporales para almacenar datos de una reserva
reserva_actual = {}
capturando_habitacion = False

# Recorrer línea por línea
for line in lines:
    line = line.strip()

    # Detectar inicio de nueva reservación
    if line.startswith("Reservacion:"):
        if reserva_actual:
            datos_procesados.append(reserva_actual)
        reserva_actual = {"Reservacion": line.split(":", 1)[1].strip()}
        capturando_habitacion = False

    # Detectar nombre
    elif line.startswith("Nombre :"):
        nombre = line.split(":", 1)[1].strip()
        partes = nombre.split()
        reserva_actual["First Name*"] = partes[0] if partes else ""
        reserva_actual["Last Name*"] = " ".join(partes[1:]) if len(partes) > 1 else ""

    # Detectar teléfono
    elif line.startswith("Telefono :"):
        reserva_actual["Phone"] = line.split(":", 1)[1].strip()

    # Detectar email
    elif line.startswith("E-mail :"):
        reserva_actual["Email*"] = line.split(":", 1)[1].strip()

    # Detectar agencia
    elif line.startswith("Agencia :"):
        reserva_actual["Source*"] = line.split(":", 1)[1].strip()

    # Detectar estado
    elif line.startswith("Status :"):
        estado = line.split(":", 1)[1].strip()
        reserva_actual["Status*"] = {
            "PENDIENTE POR LLEGAR": "Confirmed",
            "CANCELADA POR SISTEMA": "Cancelled",
            "CANCELADA POR CLIENTE": "Cancelled",
            "CANCELADA POR HOTEL": "Cancelled"
        }.get(estado, "Pending")

    # Detectar comentarios
    elif line.startswith("Comentarios:"):
        reserva_actual["Note"] = line.split(":", 1)[1].strip()

    # Detectar inicio de sección de habitación
    elif line.startswith("Clase"):
        capturando_habitacion = True

    # Capturar datos de habitación
    elif capturando_habitacion and line:
        partes = line.split(",")
        if len(partes) >= 8:
            reserva_actual["Accommodation*"] = partes[0].strip()
            fecha_entrada = partes[3].strip()
            fecha_salida = partes[4].strip()

            # Convertir manualmente de dd/mm/yyyy a yyyy-mm-dd 00:00:00 como texto
            try:
                dia, mes, anio = fecha_entrada.split("/")
                reserva_actual["Arrival Date*"] = f"{anio}-{mes.zfill(2)}-{dia.zfill(2)} 00:00:00"
            except:
                reserva_actual["Arrival Date*"] = ""

            try:
                dia, mes, anio = fecha_salida.split("/")
                reserva_actual["Departure Date*"] = f"{anio}-{mes.zfill(2)}-{dia.zfill(2)} 00:00:00"
            except:
                reserva_actual["Departure Date*"] = ""

            reserva_actual["Adults*"] = int(partes[5].strip()) if partes[5].strip().isdigit() else 1
            reserva_actual["Children*"] = int(partes[6].strip()) if partes[6].strip().isdigit() else 0
            reserva_actual["Payment amount (without taxes)"] = float(partes[8].strip()) if partes[8].strip() else 0
            capturando_habitacion = False

# Agregar la última reserva si existe
if reserva_actual:
    datos_procesados.append(reserva_actual)

# Crear DataFrame con los datos procesados
df = pd.DataFrame(datos_procesados)

# Agregar columnas faltantes del template
columnas_template = [
    'First Name*', 'Last Name*', 'Email*', 'Arrival Date*', 'Departure Date*',
    'Accommodation*', 'Phone', 'Address', 'Country*', 'State', 'Zip Code',
    'Source*', 'Status*', 'Adults*', 'Children*', 'External Reference ID',
    'Note', 'Payment Type', 'Payment amount (without taxes)'
]

for col in columnas_template:
    if col not in df.columns:
        df[col] = ''

# Reordenar columnas
df = df[columnas_template]

# Guardar a Excel
df.to_excel('reservaciones_procesadas.xlsx', index=False)

print(f"Procesamiento completado. Se han procesado {len(df)} reservaciones.")