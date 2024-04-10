import os
import re
import pandas as pd
import xlrd  # Necesario para leer archivos .xls

def find_emails_in_text(text):
    # Regex para encontrar correos electrónicos
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    return re.findall(pattern, text)

def extract_emails_from_txt(file_path):
    encodings = ['utf-8', 'latin-1', 'iso-8859-1']  # Lista de codificaciones a intentar
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as file:
                content = file.read()
            return find_emails_in_text(content)
        except UnicodeDecodeError:
            continue  # Intenta con la siguiente codificación si ocurre un error
    print(f"No se pudo leer el archivo {file_path} con ninguna de las codificaciones probadas.")
    return []

def extract_emails_from_excel(file_path):
    emails = []
    try:
        if file_path.endswith('.xlsx'):
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name)
                for column in df.columns:
                    for value in df[column]:
                        if isinstance(value, str):
                            emails.extend(find_emails_in_text(value))
        elif file_path.endswith('.xls'):
            book = xlrd.open_workbook(file_path)
            for sheet in book.sheets():
                for row in range(sheet.nrows):
                    for cell in sheet.row(row):
                        if cell.ctype == xlrd.XL_CELL_TEXT:
                            emails.extend(find_emails_in_text(cell.value))
    except Exception as e:
        print(f"Error procesando {file_path}: {e}")
    return emails

def main(folder_path):
    emails = set()
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            print(f"Procesando {file_path}...")
            if file.endswith('.txt'):
                emails.update(extract_emails_from_txt(file_path))
            elif file.endswith(('.xlsx', '.xls')):
                emails.update(extract_emails_from_excel(file_path))
    
    output_file_path = 'extracted_emails.txt'
    with open(output_file_path, 'w', encoding='utf-8') as output_file:
        for email in sorted(emails):  # Ordena los emails antes de escribirlos
            output_file.write(email + '\n')
    
    print(f"Extracción completada. Emails guardados en {output_file_path}")

if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        print("Por favor, proporciona el directorio raíz como argumento.")
