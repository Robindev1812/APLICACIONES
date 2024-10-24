import pyodbc
import pandas as pd
import os

#RUTA DE LA BD DE ACCESS

#CLASE PARA CONECTAR A LA BASE DE DATOS DE ACCESS
class AccessConnection:
    def __init__(self):
        self.db_path = r'C:\APLICACIONES - PYTHON\BASES\BD_APLICACIONES.accdb'
        self.conn_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + self.db_path
        self.conn = None
      
    #CONECTAR A LA BD  
    def connect(self):
        try:
            self.conn = pyodbc.connect(self.conn_string)
            print(f"Conexion a la base de datos exitosa")
        except Exception as e:
            print(f"Error al conectar la base de datos: {e}")

    #CREAR TABLA
    def create_table(self, table_name, columns):
        if self.conn is None:
            print("No hay conexión establecida")
            return None
        
        # Línea de SQL para crear la tabla
        columns_definition = ', '.join([f"{name} {datatype}" for name, datatype in columns.items()])
        sql_query = f"CREATE TABLE {table_name} ({columns_definition})"
        
        try:
            cursor = self.conn.cursor()
            cursor.execute(sql_query)
            self.conn.commit()  # Confirmar los cambios en la base de datos
            print(f"Tabla '{table_name}' creada exitosamente")
        except Exception as e:
            print(f"Error al crear la tabla: {e}")
        finally:
            cursor.close()  # Asegurarse de cerrar el cursor
            
    #CONSULTAR UNA TABLA DE LA BD        
    def read_table(self, name_table):
        if self.conn is None:
            print(f"No hay conexion establecida")
            return None

        try:
            df = pd.read_sql(f"SELECT * FROM {name_table}" , self.conn)
            return df
        except Exception as e:
            print(f"Error al recuperar datos...")
        

    #INSERTAR DATOS EN LAS TABLAS
    def insert_table(self, table_name, file_path):
        if self.conn is None:
            print("No hay conexión establecida")
            return
        
        try:
            with open(file_path, 'r') as file:
                #VARIABLES PARA LOS INSERT 
                NOMBRE = CEDULA = PRESTAMO = CUOTA = IMPORTE = ALTURA_CUOTA = CUOTAS_PAGADAS = PLAZO_RESTANTE = CONVENIO = CONSECUTIVO = OFICINA = SALDO_CAPITAL = PLAZO_INICIAL = NOMBRE_CONVENIO = SEGUNDA_LIB = True

                for line in file:
                    # Limpiar espacios y dividir la línea por punto y coma
                    data = line.strip().split(';')
                    
                    if len(data) < 1:  # Verificamos que hay suficientes columnas
                        print(f"Datos insuficientes en la línea: {line.strip()}")
                        continue
                    
                    # Extraemos los datos necesarios
                    if "CONTABILIZADOS":

                        NOMBRE = data[10]
                        CEDULA = data[11]
                        PRESTAMO = data[9]
                        CUOTA= data[17]
                        IMPORTE = data[13]
                        ALTURA_CUOTA= data[18]
                        CUOTAS_PAGADAS= data[19]
                        PLAZO_RESTANTE= data[20]
                        CONVENIO= data[1]
                        CONSECUTIVO= data[3]
                        OFICINA= data[8]
                        SALDO_CAPITAL= data[23]
                        PLAZO_INICIAL= data[12]
                        NOMBRE_CONVENIO= data[2]
                        SEGUNDA_LIB = ""
                    
                    # Crear la consulta SQL para insertar los datos
                    sql_query = f"INSERT INTO {table_name} (NOMBRE, CEDULA, PRESTAMO, CUOTA, IMPORTE, ALTURA_CUOTA, CUOTAS_PAGADAS, PLAZO_RESTANTE, CONVENIO, CONSECUTIVO, OFICINA, SALDO_CAPITAL, PLAZO_INICIAL, NOMBRE_CONVENIO, SEGUNDA_LIB) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                    
                    # Ejecutar la consulta con los datos
                    cursor = self.conn.cursor()
                    cursor.execute(sql_query, (NOMBRE, CEDULA, PRESTAMO, CUOTA, IMPORTE, ALTURA_CUOTA, CUOTAS_PAGADAS, PLAZO_RESTANTE, CONVENIO, CONSECUTIVO, OFICINA, SALDO_CAPITAL, PLAZO_INICIAL, NOMBRE_CONVENIO, SEGUNDA_LIB))
                    
            self.conn.commit()  # Confirmar los cambios en la base de datos
            print("Datos insertados exitosamente")
        except Exception as e:
            print(f"Error al insertar los datos: {e}")
        finally:
            cursor.close()  # Asegurarse de cerrar el cursor


    #ACTUALIDAR UNA TABLA
    def update_table(self, name_table):
        pass
    
    def delete_table(self, name_table):
        if self.conn is None:
            print("No hay conexión establecida")
            return None

        sql_query = f"DELETE * FROM {name_table}"

        try:
            cursor = self.conn.cursor()
            cursor.execute(sql_query)
            self.conn.commit()  # Confirmar los cambios en la base de datos
            print(f"Informacion eliminada de la tabla {name_table}")
        except Exception as e:
            print(f"Error al eliminar tabla: {e}")
        finally:
            cursor.close()  # Asegurarse de cerrar el cursor

        

    #CERRAR LA CONEXÍON
    def close_connection(self):
        if self.conn is not None:
            self.conn.close()
            print("Conexion cerrada")


        
        
                    
        
