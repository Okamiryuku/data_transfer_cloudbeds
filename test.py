import pandas as pd

# Leer sin encabezados
df = pd.read_excel('Reservaciones.xlsx', header=None)

def extract_reservation_data(df):
    reservations = []
    current_reservation = {}

    for idx, row in df.iterrows():
        cell_0 = str(row[0]) if pd.notna(row[0]) else ''

        # Nueva reservación
        if 'Reservacion:' in cell_0:
            if current_reservation:
                reservations.append(current_reservation)
            current_reservation = {
                'External Reference ID': str(row[1]).split()[0] if pd.notna(row[1]) else '',
                'First Name': '',
                'Last Name': '',
                'Email': '',
                'Phone': '',
                'Arrival Date': '',
                'Departure Date': '',
                'Accommodation': '',
                'Adults': '',
                'Children': '',
                'Country': '',
                'State': '',
                'Zip Code': '',
                'Address': '',
                'Source': '',
                'Status': '',
                'Payment Type': '',
                'Payment amount (without taxes)': '',
                'Note': ''
            }

        # Nombre
        elif 'Nombre :' in cell_0:
            full_name = str(row[1]).strip()
            parts = full_name.split(maxsplit=1)
            current_reservation['First Name'] = parts[0] if parts else ''
            current_reservation['Last Name'] = parts[1] if len(parts) > 1 else ''

        # Email
        elif 'E-mail :' in cell_0:
            current_reservation['Email'] = str(row[7]).strip() if pd.notna(row[7]) else ''

        # Teléfono
        elif 'Telefono :' in cell_0:
            current_reservation['Phone'] = str(row[1]).strip() if pd.notna(row[1]) else ''

        # --- CORRECCIÓN: detectar fila con "F. Entrada" en cualquier columna ---
        elif any('F. Entrada' in str(c) for c in row):
            try:
                data_row = df.iloc[idx + 2]
                # Fecha de entrada (columna 3)
                arr_raw = data_row.iloc[3] if pd.notna(data_row.iloc[3]) else None
                if arr_raw:
                    arr_dt = pd.to_datetime(arr_raw, format='%d/%m/%Y', errors='coerce')
                    current_reservation['Arrival Date'] = arr_dt.strftime('%Y-%m-%d') if pd.notna(arr_dt) else ''
                # Fecha de salida (columna 5)
                dep_raw = data_row.iloc[4] if pd.notna(data_row.iloc[4]) else None
                if dep_raw:
                    dep_dt = pd.to_datetime(dep_raw, format='%d/%m/%Y', errors='coerce')
                    current_reservation['Departure Date'] = dep_dt.strftime('%Y-%m-%d') if pd.notna(dep_dt) else ''
                # Otros campos de la misma fila
                current_reservation['Accommodation'] = str(data_row.iloc[0]).strip() if pd.notna(data_row.iloc[0]) else ''
                current_reservation['Adults'] = str(data_row.iloc[5]).strip() if pd.notna(data_row.iloc[5]) else '0'
                current_reservation['Children'] = str(data_row.iloc[6]).strip() if pd.notna(data_row.iloc[6]) else '0'
                current_reservation['Payment amount (without taxes)'] = str(data_row.iloc[9]).strip() if pd.notna(data_row.iloc[9]) else ''
            except IndexError:
                pass

        # Procedencia
        elif 'Procedencia :' in cell_0:
            procedencia = str(row[7]).strip() if pd.notna(row[7]) else ''
            if ',' in procedencia:
                estado, pais = procedencia.split(',', 1)
                current_reservation['State'] = estado.strip()
                current_reservation['Country'] = pais.strip()
            else:
                current_reservation['Country'] = procedencia

        # Estado
        elif 'Status :' in cell_0:
            status = str(row[1]).strip()
            current_reservation['Status'] = 'Confirmed' if 'PENDIENTE POR LLEGAR' in status else 'Cancelled' if 'CANCELADA' in status else status

        # Agencia
        elif 'Agencia :' in cell_0:
            agencia = str(row[1]).strip()
            mapping = {
                'BOOKING': 'Booking.com',
                'EXPEDIA': 'Expedia',
                'ROIBACK': 'Roiback',
                'PRESENCIAL': 'Direct',
                'TELEFONICA': 'Phone',
                'WHATSAPP': 'WhatsApp',
                'CORREO ELECTRÓNICO': 'Email'
            }
            current_reservation['Source'] = mapping.get(agencia, agencia)

        # Garantía / forma de pago
        elif 'Garantia :' in cell_0:
            garantia = str(row[7]).strip() if pd.notna(row[7]) else ''
            mapping = {
                'PREPAGO': 'Prepaid',
                'AUT PTE': 'Credit Card',
                'TRANSFERENCIA': 'Bank Transfer',
                'TB': 'Bank Transfer',
                'LLIGA DE PAGO DE CLIP': 'Clip Payment'
            }
            current_reservation['Payment Type'] = mapping.get(garantia, garantia)

        # Notas
        elif 'Comentarios:' in cell_0:
            current_reservation['Note'] = str(row[1]).strip() if pd.notna(row[1]) else ''

    if current_reservation:
        reservations.append(current_reservation)
    return reservations

# Procesar
reservas = extract_reservation_data(df)

# Columnas finales según el template
cols = [
    'First Name*', 'Last Name*', 'Email*', 'Arrival Date*', 'Departure Date*',
    'Accommodation*', 'Phone', 'Address', 'Country*', 'State', 'Zip Code',
    'Source*', 'Status*', 'Adults*', 'Children*', 'External Reference ID',
    'Note', 'Payment Type', 'Payment amount (without taxes)'
]

df_out = pd.DataFrame(reservas).rename(columns={
    'First Name': 'First Name*',
    'Last Name': 'Last Name*',
    'Email': 'Email*',
    'Arrival Date': 'Arrival Date*',
    'Departure Date': 'Departure Date*',
    'Accommodation': 'Accommodation*',
    'Country': 'Country*',
    'Source': 'Source*',
    'Status': 'Status*',
    'Adults': 'Adults*',
    'Children': 'Children*'
})[cols]

df_out.to_excel('Reservaciones_Formateadas.xlsx', index=False)
print('Archivo generado: Reservaciones_Formateadas.xlsx')