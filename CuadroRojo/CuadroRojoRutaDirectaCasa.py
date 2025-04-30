import pandas as pd
import time
import openpyxl
from pathlib import Path
import threading

# --- Función para mostrar el tiempo transcurrido ---
def mostrar_tiempo():
    inicio = time.time()
    while True:
        transcurrido = int(time.time() - inicio)
        minutos = transcurrido // 60
        segundos = transcurrido % 60
        print(f"\r⏳ Tiempo Transcurrido: {minutos:02d}:{segundos:02d}", end="")
        time.sleep(1)

# --- Cargar datos principales ---
ruta_origen = Path(r'C:\Users\MARCOS  MARQUEZ\Desktop\AIR-E\Automatizaciones\CuadroRojo\CarpetasBase\Resumen Estad Proy. Viab.xlsx')
ruta_destino = Path(r'C:\Users\MARCOS  MARQUEZ\Desktop\AIR-E\Automatizaciones\CuadroRojo\CarpetasBase\Análisis Generación_VF.xlsx')
ruta_base = Path(r'C:\Users\MARCOS  MARQUEZ\Desktop\AIR-E\Automatizaciones\CuadroRojo\CarpetasBase')

excel = pd.read_excel(ruta_origen, header=3)  # Datos principales
excel_destino = pd.read_excel(ruta_destino, header=5, dtype=str)

# Columnas de interés
col_nombre_proyecto = excel['NOMBRE PROYECTO']
col_fecha_fin = excel['FEC FIN']
col_prst = excel['PRST']
col_estado_inicial = excel['ESTADO INICIAL']
col_otaire = excel['OT AIR-E']
cantidad_proyectos = len(col_nombre_proyecto)

carpetas_prst = [c.name for c in ruta_base.iterdir() if c.is_dir()]

# --- Procesamiento de proyectos ---
for i in range(cantidad_proyectos):
    try:
        # NO convertir a minúsculas aquí
        nombre_proyecto = col_nombre_proyecto[i].strip()
        nombre_prst = col_prst[i].strip()
        estado_inicial = col_estado_inicial[i].strip()

        print(f"\n🔎 Procesando: {nombre_proyecto}")

        # Filtrar solo los de "gestionar in situ"
        if estado_inicial.lower() != "gestionar in situ":
            print(f"🚫 Estado inicial '{estado_inicial}' no es 'gestionar in situ', se omite.")
            continue

        # Buscar carpeta PRST (respetando mayúsculas/minúsculas)
        carpeta_prst = next((c for c in carpetas_prst if c.strip() == nombre_prst), None)
        if not carpeta_prst:
            print(f"❌ No se encontró la carpeta PRST: {nombre_prst}")
            continue

        ruta_prst = ruta_base / carpeta_prst
        carpetas_proyecto = [c.name for c in ruta_prst.iterdir() if c.is_dir()]

        # Buscar carpeta Proyecto
        carpeta_proyecto = next((c for c in carpetas_proyecto if c.strip() == nombre_proyecto), None)
        if not carpeta_proyecto:
            print(f"❌ No se encontró la carpeta del Proyecto: {nombre_proyecto}")
            continue

        ruta_proyecto = ruta_prst / carpeta_proyecto
        carpeta_gestion = ruta_proyecto / '2. Gestion'

        if not carpeta_gestion.exists():
            print(f"⚠️ No existe carpeta '2. Gestion' en {ruta_proyecto}")
            continue

        # Buscar archivo
        archivo_excel = carpeta_gestion / 'Inf_Viabilidad_In_Situ.xlsx'

        if not archivo_excel.exists():
            print(f"❌ Archivo no encontrado: {archivo_excel.name}")
            continue

        print(f"📄 Archivo encontrado: {archivo_excel.name}")

        # Leer archivo inspección
        try:
            df_inspeccion = pd.read_excel(archivo_excel, sheet_name='INSPECCIÓN', header=1)
            df_viabilidad = pd.read_excel(archivo_excel, sheet_name='VIABILIDAD', header=5)
            df_viabilidad_raw = pd.read_excel(archivo_excel, sheet_name='VIABILIDAD', header=None)
        except Exception as e:
            print(f"❌ Error leyendo hojas de Excel: {e}")
            continue

        # Leer nombre PRST largo
        nombre_prst_largo = df_viabilidad_raw.iloc[1, 2] if not df_viabilidad_raw.empty else ''

        # Procesar filas
        for idx, fila in df_inspeccion.iterrows():
            # Buscar filas vacías en el archivo destino
            filas_vacias = excel_destino[excel_destino['Fecha de Brigada'].isna()]
            if filas_vacias.empty:
                print("🚫 No hay filas vacías disponibles en el destino.")
                break

            index_destino = filas_vacias.index[0]

            # Mapeo campos
            mappings = {
                'Fecha de Brigada': fila.get('Fecha', ''),
                'Número del Poste': fila.get('N° Poste en plano', ''),
                'Elementos Existentes': fila.get('Elementos Existentes', ''),
                'Material del Poste': fila.get('Material', ''),
                'Tipo de Poste': fila.get('Tipo de poste', ''),
                'Altura del Poste': fila.get('Altura', ''),
                'Municipio': excel.loc[i, 'MUNICIPIO'],
                'Departamento': excel.loc[i, 'DEPARTAMENTO'],
                'Cantidad de PRST en Poste': fila.get('Cantidad de PRST en poste', ''),
                'Nombre de PRST - Largo': nombre_prst_largo,
                'Nombre de PRST - Corto': col_prst[i],
                'Fecha de Alta': col_fecha_fin[i],
                'Tipo de Coordenada': "Real",
                'Nombre del Proyecto': col_nombre_proyecto[i],
                'OT AIRE': col_otaire[i],
                'Observaciones': fila.get('Observaciones', ''),
                'RF1': fila.get('RF_1', ''),
                'RF2': fila.get('RF_2', ''),
                'RF3': fila.get('RF_3', ''),
                'Coordenada_X': fila.get('point_x', ''),
                'Coordenada_Y': fila.get('point_y', '')
            }

            # Valor especial: 'Cantidad' desde hoja Viabilidad
            mappings['Cantidad'] = df_viabilidad['Cables '].iloc[0] if 'Cables ' in df_viabilidad.columns else ''

            # Asignar datos a la fila vacía
            for col, val in mappings.items():
                excel_destino.at[index_destino, col] = val

        print(f"✅ Proyecto {col_nombre_proyecto[i]} procesado correctamente.")

    except Exception as e:
        print(f"❌ Error general procesando el proyecto {col_nombre_proyecto[i]}: {e}")

# --- Guardar cambios SOLO las filas de datos ---
try:
    # Cargar el workbook
    wb = openpyxl.load_workbook(ruta_destino)
    ws = wb.active  # Si tienes varias hojas, especifica el nombre

    for idx, row in excel_destino.iterrows():
        # Nos interesa solo a partir de la fila 6 en excel (índice 5 en pandas)
        excel_row = idx + 1  # Excel empieza en 1, pandas en 0
        if excel_row < 6:  # Evitar las filas de encabezados
            continue

        for col_idx, value in enumerate(row, 1):  # Columnas empiezan en 1
            if pd.notna(value):
                ws.cell(row=excel_row + 1, column=col_idx, value=value)

    wb.save(ruta_destino)
    print("\n✅ Archivo destino actualizado correctamente.")
except Exception as e:
    print(f"❌ Error guardando el archivo destino: {e}")

# --- Mostrar tiempo transcurrido ---
mostrar_tiempo()
print("\n⏳ Proceso finalizado.")
