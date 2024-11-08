from concurrent.futures import ThreadPoolExecutor
from PyQt5 import QtWidgets
import pandas as pd
import time
import sys

class ProcesarPanda:
    hoja_col5 = "operaciones"
    hoja_col10 = "materiales"
    ruta_excel = "./datos.xlsx"

    def validar_estructura_general(self):
        start_time = time.time()  # Inicia el contador de tiempo
        try:
            with pd.ExcelFile(self.ruta_excel) as excel:
                # Leer las hojas necesarias
                df_operaciones = excel.parse(self.hoja_col5)
                df_materiales = excel.parse(self.hoja_col10)

                # Verificar número mínimo de columnas
                if df_operaciones.shape[1] < 10 or df_materiales.shape[1] < 10:
                    return False

                # Validar el formato de los datos
                df_col5 = df_operaciones.iloc[:, 4]  # Columna 5 (índice 4)
                df_col10 = df_materiales.iloc[:, 9]  # Columna 10 (índice 9)
                pattern = r'^\d{4}-[A-Z]{3}$'

                if not df_col5.str.match(pattern).all() or not df_col10.str.match(pattern).all():
                    return False

            # Mostrar tiempo de validación
            elapsed_time = time.time() - start_time
            print ("Pandas")
            print(f"Tiempo de validación de estructura: {elapsed_time:.4f} segundos")
            return True

        except Exception as e:
            self.mostrar_error(f"Error al validar la estructura del archivo Excel: {e}")
            return False

    def procesar_columna(self, numero_columna):
        start_time = time.time()
        try:
            # Determinar hoja y columnas a utilizar
            if numero_columna == 4:
                hoja = self.hoja_col5
                usecols = [4]

            elif numero_columna == 9:
                hoja = self.hoja_col10
                usecols = [9]
            else:
                raise ValueError("Número de columna no válido. Debe ser 4 o 9.")

            # Leer la hoja y columna especificada
            df = pd.read_excel(self.ruta_excel, sheet_name=hoja, usecols=usecols)

            # Eliminar duplicados solo para la columna 4
            if numero_columna == 4:
                df = df.drop_duplicates()

            print(f"Tiempo para procesar columna {numero_columna}: {time.time() - start_time:.4f} segundos")
            return df

        except Exception as e:
            self.mostrar_error(f"Error al procesar la columna {numero_columna}: {e}")
            return None

    def repetidos(self):
        start_time = time.time()
        try:
            # Validar la estructura general
            with ThreadPoolExecutor() as executor:
                validacion_future = executor.submit(self.validar_estructura_general)
                es_valido = validacion_future.result()

            if not es_valido:
                self.mostrar_error("La estructura del archivo Excel no es válida.")
                return None

            # Procesar columnas en paralelo
            with ThreadPoolExecutor() as executor:
                futuro_col5 = executor.submit(self.procesar_columna, 4)
                futuro_col10 = executor.submit(self.procesar_columna, 9)

                # Obtener resultados
                valores_col5 = futuro_col5.result()
                valores_col10 = futuro_col10.result()

            if valores_col5 is None or valores_col10 is None:
                self.mostrar_error("Uno de los DataFrames es None. Verifica la función de procesamiento.")
                return None

            # Convertir DataFrames a listas
            valores_col5_list = valores_col5.iloc[:, 0].tolist()  # Columna de únicos
            valores_col10_list = valores_col10.iloc[:, 0].tolist()  # Columna de todos los valores

            # Contar ocurrencias de valores únicos de col5 en col10
            conteo = {valor: valores_col10_list.count(valor) for valor in valores_col5_list}
            print(f"Tiempo para contar números repetidos: {time.time() - start_time:.4f} segundos")
            return conteo

        except Exception as e:
            self.mostrar_error(f"Error al hacer recuento de números repetidos: {e}")
            return None

    def mostrar_error(self, mensaje):
        # Crear una aplicación Qt solo si no existe
        app = QtWidgets.QApplication.instance()
        if app is None:
            app = QtWidgets.QApplication(sys.argv)

        # Crear un cuadro de mensaje de error
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setWindowTitle("Error")
        msg.setText(mensaje)
        msg.exec_()
