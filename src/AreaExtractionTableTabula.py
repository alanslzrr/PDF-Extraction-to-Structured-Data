import matplotlib
import camelot
import matplotlib.pyplot as plt
matplotlib.use('TkAgg')  # Usa el backend de TkAgg, o 'Qt5Agg'



tables = camelot.read_pdf('01-pdf-to-xlsx/pdf_files/C295447.pdf', pages='2', flavor='stream')
camelot.plot(tables[0], kind='text').show()
plt.show()

# Definimos el área de la tabla excluyendo encabezados y pies de página (coordenadas x1, y1, x2, y2)
# Las coordenadas deben ser ajustadas específicamente para la estructura de pdf, en este caso, concluimos con table_areas = ["0,700,600,50"]
