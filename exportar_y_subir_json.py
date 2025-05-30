import pandas as pd
import subprocess
import webbrowser

EXCEL_FILE = "ACTIVOS VISA.xlsx"
SHEET_NAME = "Combinar2"
JSON_FILE = "activos_visa.json"

COLUMNAS = [
    "ID", "COD_ORG", "DESCRIPCION", "MODELO", "NUMERO SERIE",
    "FECHA_ALTA", "FECHA_PLANIFICADA", "DEPARTAMENTO"
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
    for campo_fecha in ["FECHA_ALTA", "FECHA_PLANIFICADA"]:
        if campo_fecha in df.columns:
            df[campo_fecha] = pd.to_datetime(df[campo_fecha], errors='coerce').dt.strftime('%Y-%m-%d')

    print("📤 Exportando a JSON...")
    df.to_json(JSON_FILE, orient='records', indent=4, force_ascii=False)

def subir_a_github():
    print("🔄 Preparando para subir a GitHub...")

    # Solo añadimos el JSON al área de staging
    subprocess.run(["git", "add", JSON_FILE], check=True)

    # Guardamos el commit si hay cambios
    result = subprocess.run(["git", "diff", "--cached", "--quiet"])
    if result.returncode == 0:
        print("ℹ️ No hay cambios nuevos.")
        return

    subprocess.run(["git", "commit", "-m", "Actualización automática de activos_visa.json"], check=True)

    # Descartamos temporalmente los cambios sin guardar (como el Excel)
    subprocess.run(["git", "stash", "--keep-index"], check=True)

    # Hacemos pull y push
    subprocess.run(["git", "pull", "--rebase"], check=True)
    subprocess.run(["git", "push"], check=True)

    # Recuperamos los cambios no subidos (como el Excel)
    subprocess.run(["git", "stash", "pop"], check=True)

    print("✅ JSON subido correctamente.")

if __name__ == "__main__":
    exportar_a_json()
    subir_a_github()
