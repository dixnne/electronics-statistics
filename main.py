import tabula

# Extract tables from the PDF
tables = tabula.read_pdf('Tarea 1.2 Lecturas del Problema 2025.pdf', pages='all')

# Print the number of tables extracted
print(f"Number of tables extracted: {len(tables)}")

# Print the first table
print(tables[0])