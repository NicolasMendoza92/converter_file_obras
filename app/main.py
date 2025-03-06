from fastapi import FastAPI, UploadFile, HTTPException, File
import pandas as pd
from io import BytesIO
from fastapi.middleware.cors import CORSMiddleware
import json

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

last_json_output = None

@app.get("/")
def read_root():
    if last_json_output:
        return { "data": last_json_output }
    return {"message": "Subi el archivo mada faka"}

def handle_upload_file(xls: pd.ExcelFile, sheet_name: str) -> str:
    try:
        # Leer la hoja del archivo sin encabezados iniciales
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        # Buscar la fila donde están los encabezados correctos
        header_row = None
        for idx, row in df.iterrows():
            if "Nº" in row.values and "CONCEPTO" in row.values:
                header_row = idx
                break

        if header_row is None:
            raise ValueError("No se encontraron los encabezados esperados en la hoja de Excel.")

        # Leer los datos a partir de la fila correcta con los nombres de columna adecuados
        df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_row)

        # Filtrar solo las columnas necesarias
        columnas_necesarias = ["Nº", "CONCEPTO", "UT", "CANT.", "PRECIO UNIT."]
        df = df[columnas_necesarias]

        # Filtrar filas con el patrón deseado en la columna "Nº"
        df = df[df["Nº"].astype(str).str.match(r"^[A-Z]\.?\d*$", na=False)]

        # Detectar la última fila válida
        ultima_fila_valida = df.index[-1] if not df.empty else None
        if ultima_fila_valida and ultima_fila_valida + 1 < len(df):
            print("⚠️ Advertencia: Hay datos adicionales después de la última fila válida.")

        # Estructurar los datos en una lista
        structured_data = []
        current_section = None

        for _, row in df.iterrows():
            num = str(row["Nº"]).strip() if pd.notna(row["Nº"]) else ""
            concepto = str(row["CONCEPTO"]).strip() if pd.notna(row["CONCEPTO"]) else ""
            unidad = row["UT"] if pd.notna(row["UT"]) else ""
            cantidad = row["CANT."] if pd.notna(row["CANT."]) else 0.0
            precio = row["PRECIO UNIT."] if pd.notna(row["PRECIO UNIT."]) else 0.0

            # Si "Nº" tiene solo una letra (ej: "A"), es una sección
            if len(num) == 1 and num.isalpha():
                current_section = concepto
            # Si "Nº" tiene una letra seguida de un número (ej: "A.1"), es un ítem
            elif "." in num and current_section:
                structured_data.append({
                    "section": current_section,
                    "description": concepto,
                    "unit": unidad,
                    "quantity": round(cantidad, 2),
                    "price": round(precio, 2)
                })
                
        # Creo el DF intermedio, por que queiro limpiarlo un poco mas
        df_intermedio = pd.DataFrame(structured_data)
        print(df_intermedio.head(10))  
        # Redondear el precio a 2 decimales
        df_intermedio["price"] = df_intermedio["price"].round(2)
        df_intermedio["quantity"] = df_intermedio["quantity"].round(2)
        
        json_output = json.loads(df_intermedio.to_json(orient="records", force_ascii=False))
        # Convertir a JSON
        return  json_output

    except Exception as e:
        raise ValueError(f"Error procesando los datos: {str(e)}")

@app.post("/convert")
async def upload_excel(file: UploadFile = File(...)):
    try:
        # Leer el contenido del archivo
        contents = await file.read()
        xls = pd.ExcelFile(BytesIO(contents))

        # Validar que el archivo tiene hojas
        if not xls.sheet_names:
            raise HTTPException(status_code=400, detail="El archivo Excel no contiene hojas.")

        # Tomar la primera hoja
        sheet_name = xls.sheet_names[0]

        # Llamar a la función de procesamiento
        last_json_output = handle_upload_file(xls, sheet_name)

        return {"data": last_json_output}

    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error procesando el archivo: {str(e)}")