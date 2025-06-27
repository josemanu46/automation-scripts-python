import os

def clean_excel_files(directory_path):
    if not os.path.exists(directory_path):
        print(f"La carpeta '{directory_path}' no existe.")
        return

    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        try:
            if os.path.isfile(file_path) and file_path.endswith('.xlsx') or file_path.endswith('.xlsm'):
                os.unlink(file_path)  # Elimina el archivo .xlsx
                print(f"Archivo eliminado: {file_path}")
        except Exception as e:
            print(f"No se pudo eliminar {file_path}. Razón: {e}")

    print(f"Todos los archivos .xlsx en la carpeta '{directory_path}' han sido eliminados.")
