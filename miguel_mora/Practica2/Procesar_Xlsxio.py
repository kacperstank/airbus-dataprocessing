import re
from concurrent.futures import ThreadPoolExecutor
import time

import xlsxio


class Procesar_Xlsxio:
    hoja_col5 = "operaciones"
    hoja_col10 = "materiales"
    ruta_excel = "./datos.xlsx"

    def __init__(self):
        self.pattern = re.compile(r'^\d{4}-[A-Z]{3}$')

    def validar_estructura(self):
        start_time = time.time()
        try:
            with xlsxio.XlsxioReader(self.ruta_excel) as reader:
                # Validación para hoja de operaciones (col5)
                with reader.get_sheet(self.hoja_col5) as sheet:
                    header = sheet.read_header()
                    if len(header) < 10:
                        return False
                    for row in sheet.iter_rows():
                        if not self.pattern.match(str(row[4])):
                            return False

                # Validación para hoja de materiales (col10)
                with reader.get_sheet(self.hoja_col10) as sheet:
                    header = sheet.read_header()
                    if len(header) < 10:
                        return False
                    for row in sheet.iter_rows():
                        if not self.pattern.match(str(row[9])):
                            return False
            print ("XLSXIO")
            print(f"Tiempo de validación de estructura: {time.time() - start_time:.4f} segundos")
            return True
        except Exception as e:
            print(f"Error al validar estructura: {e}")
            return False

    def procesar_columna(self, numero_columna):
        start_time = time.time()
        datos = set() if numero_columna == 4 else []  # Conjunto para únicos, lista para todos los valores
        try:
            hoja = self.hoja_col5 if numero_columna == 4 else self.hoja_col10
            with xlsxio.XlsxioReader(self.ruta_excel) as reader:
                with reader.get_sheet(hoja) as sheet:
                    for row in sheet.iter_rows():
                        if numero_columna < len(row):
                            if numero_columna == 4:
                                datos.add(row[numero_columna])  # Agregar único al conjunto
                            else:
                                datos.append(row[numero_columna])  # Agregar todos los valores a la lista

            print(f"Tiempo de procesamiento de columna {numero_columna}: {time.time() - start_time:.4f} segundos")
            return list(datos) if isinstance(datos, set) else datos  # Convertir conjunto a lista si es necesario
        except Exception as e:
            print(f"Error al procesar los datos de la columna {numero_columna}: {e}")
            return None

    def repetidos(self):
        start_time = time.time()
        try:
            with ThreadPoolExecutor() as executor:
                validacion_future = executor.submit(self.validar_estructura)
                es_valido = validacion_future.result()

            if not es_valido:
                print("La estructura del archivo Excel no es válida.")
                return None

            with ThreadPoolExecutor() as executor:
                futuro_col5 = executor.submit(self.procesar_columna, 4)
                futuro_col10 = executor.submit(self.procesar_columna, 9)

                valores_col5 = futuro_col5.result()
                valores_col10 = futuro_col10.result()

            if valores_col5 is None or valores_col10 is None:
                print("Uno de los DataFrames es None. Verifica la función de procesamiento.")
                return None

            conteo = {valor: valores_col10.count(valor) for valor in valores_col5}
            print(f"Tiempo para contar números repetidos: {time.time() - start_time:.4f} segundos")
            return conteo
        except Exception as e:
            print(f"Error al hacer recuento de números repetidos: {e}")
            return None
