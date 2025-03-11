from fastapi import FastAPI, UploadFile, HTTPException, File, Response
import pandas as pd
from io import BytesIO
from fastapi.middleware.cors import CORSMiddleware
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from datetime import datetime
from typing import Dict, Any
import json
import locale

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

        return last_json_output

    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error procesando el archivo: {str(e)}")
    
try:
    locale.setlocale(locale.LC_ALL, locale.getdefaultlocale()) 
except locale.Error:
    locale.setlocale(locale.LC_ALL, '')

def generate_pdf_format(data: Dict[str, Any], pdf_buffer: BytesIO):
    c = canvas.Canvas(pdf_buffer, pagesize=letter)
    project_data = data['Project']
    certificate_data = data

    # Extraer datos del proyecto
    project_name = project_data['name']
    project_number = project_data['projectNumber']
    project_address = project_data['address']
    project_description = project_data['description']
    issued_at = datetime.strptime(certificate_data['issuedAt'], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d/%m/%Y")

    # Título
    c.setFont('Helvetica', 22)
    c.drawString(30, 750, 'Certificado de Avance de Obra')
    c.setFont('Helvetica', 12)
    c.drawString(30, 735, f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Detalles del proyecto
    c.setFont('Helvetica-Bold', 14)
    c.drawString(30, 700, "Detalles del Proyecto:")
    c.setFont('Helvetica', 12)
    c.drawString(30, 680, f"Nombre: {project_name}")
    c.drawString(30, 660, f"Número de Proyecto: {project_number}")
    c.drawString(30, 640, f"Dirección: {project_address}")
    c.drawString(30, 620, f"Descripción: {project_description}")

    # Detalles del Certificado
    c.setFont('Helvetica-Bold', 14)
    c.drawString(30, 590, "Detalles del Certificado:")
    c.setFont('Helvetica', 12)
    c.drawString(30, 570, f"Versión: {certificate_data['version']}")
    c.drawString(
        30, 550, f"Monto Certificado: {locale.currency(certificate_data['certificateAmount'], grouping=True)}"
    )
    c.drawString(30, 530, f"Fecha de Emisión: {issued_at}")

    # Tabla de Items
    data_items = [
        ["Sección", "Descripción", "Un.", "Cant.", "Precio", "Progreso", "Subtotal"]
    ]
    total_amount = 0.0

    for item_data in certificate_data['certificateItems']:
        item = item_data['item']
        progress = item_data['progress']
        subtotal = (item['price'] * progress) / 100
        total_amount += subtotal

        description = item['description'].strip()
        section = item['section'].capitalize()

        # Formateo de sección en saltos de línea
        section_lines = [section[i:i+15] for i in range(0, len(section), 15)]
        section = "\n".join(section_lines)

        # Formateo de descripción en saltos de línea
        words = description.split()
        description_lines = []
        line = ""

        for word in words:
            if len(line) + len(word) < 40:  
                line += word + " "
            else:
                description_lines.append(line.strip())
                line = word + " "
        description_lines.append(line.strip())

        data_items.append([
            section,
            "\n".join(description_lines),  
            item['unit'],
            str(item['quantity']),
            locale.currency(item['price'], grouping=True),
            f"{progress}%",
            locale.currency(subtotal, grouping=True),
        ])

    # Estilo de la tabla
    table = Table(data_items, colWidths=[80, 200, 30, 30, 70, 50, 70])
    table.setStyle(
        TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (0, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ])
    )

    table.wrapOn(c, 800, 600)
    table.drawOn(c, 30, 300 - len(data_items) * 15) 

    # Total general
    c.setFont('Helvetica-Bold', 14)
    c.drawString(30, 250 - len(data_items) * 15, f"Total del Certificado: {locale.currency(total_amount, grouping=True)}")

    c.save()
    pdf_buffer.seek(0)
    

@app.post("/generate_certificate_pdf")
async def generate_certificate_pdf(data: Dict[str, Any]):
    try:
        pdf_buffer = BytesIO()  
        generate_pdf_format(data, pdf_buffer)
        return Response(pdf_buffer.read(), media_type="application/pdf", headers={
            "Content-Disposition": "attachment; filename=certificado.pdf"
        })
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))