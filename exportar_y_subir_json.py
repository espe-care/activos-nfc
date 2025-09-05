import pandas as pd

EXCEL_FILE = "ACTIVOS TODOS1.xlsx"
SHEET_NAME = "AUX_DIM_ACTIVOS"
JSON_FILE = "activos_visa.json"

COLUMNAS = [
    "ID", "COD_ORG", "DESCRIPCION", "MODELO", "NUMERO SERIE",
    "FECHA ALTA", "FECHA_PLANIFICADA", "DEPARTAMENTO","FECHA ULTIMO RPEVENTIVO"
]

def exportar_a_json():
    print("📥 Leyendo Excel...")
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str)
    df.columns = df.columns.str.strip()

    # Filtrar solo columnas necesarias y existentes
    columnas_presentes = [col for col in COLUMNAS if col in df.columns]
    df = df[columnas_presentes]

    # Renombrar columnas para estandarizar nombres con guiones bajos
    df = df.rename(columns={
        "FECHA ALTA": "FECHA_ALTA"
    })

    # Formatear fechas a YYYY-MM-DD sin hora
    for campo_fecha in ["FECHA_ALTA", "FECHA_PLANIFICADA", "FECHA ULTIMO RPEVENTIVO"]:
        if campo_fecha in df.columns:
            df[campo_fecha] = pd.to_datetime(df[campo_fecha], errors='coerce').dt.strftime('%Y-%m-%d')

    print("📤 Exportando a JSON...")
    df.to_json(JSON_FILE, orient='records', indent=4, force_ascii=False)
    print(f"✅ Archivo JSON generado: {JSON_FILE}")

if __name__ == "__main__":
    exportar_a_json()
