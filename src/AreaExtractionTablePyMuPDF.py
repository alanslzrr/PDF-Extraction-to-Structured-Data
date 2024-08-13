import fitz  # PyMuPDF

# Ruta del archivo PDF
pdf_path = 'C295447.pdf'

# Abrir el documento PDF
doc = fitz.open(pdf_path)

# Seleccionar la primera página
page = doc[0]

# Extraer la información del texto en la página, incluyendo coordenadas
text_instances = page.get_text("dict")["blocks"]

for instance in text_instances:
    if "lines" in instance:  # Asegurarse de que es un bloque de texto
        for line in instance["lines"]:
            for span in line["spans"]:
                print(f"Texto: {span['text']} - Coordenadas: {span['bbox']}")
                
# Cerrar el documento
doc.close()
