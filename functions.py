import db
import os
import win32com.client

def create_db():
  db_path = r'C:\APLICACIONES - PYTHON\BASES\BD_APLICACIONES.accdb'

  if os.path.exists(db_path):
      print(f"LA BASE DE DATOS SE ENCUENTRA CREADA {db_path}")
      return True
  else:
      #crea la base de datos
      try:
          # Inicializar el objeto de Access
          access_app = win32com.client.Dispatch('Access.Application')
          
          #SE CREA LA BASE DE DATOS
          access_app.NewCurrentDatabase(db_path)
          print(f"LA BASE DE DATOS HA SIDO CREADA EN {db_path}")
          access_app.Quit()

          #SE CREAN LAS TABLAS DE LA BASE DE DATOS
          conn = db.AccessConnection()
          conn.connect()

          #SE CREA LA TABLA CONTABILIZADOS
          columna_contabilizados = colums_contabilizado()
          conn.create_table("CONTABILIZADOS",columna_contabilizados)

          #SE CREA LA TABLA PARAMETRÍA
          columna_parametria = colums_parametria()
          conn.create_table("PARAMETRIA",columna_parametria)

          #SE CREA LA TABLA PRÉSTAMO
          columna_prestamo = colums_prestamo()
          conn.create_table("PRESTAMO",columna_prestamo)

          #SE CREA LA TABLA RESIDUALES
          columna_residuales = colums_prestamo()
          conn.create_table("RESIDUALES",columna_prestamo)

          #SE CREA LA TABLA SOBRANTES
          columna_sobrantes = colums_sobrante()
          conn.create_table("SOBRANTES",columna_sobrantes)

          #SE CREA LA TABLA INSOLVENCIAS
          columna_insolvencias = colums_sobrante()
          conn.create_table("INSOLVENCIAS",columna_insolvencias)

      except Exception as e:
          print(f"ERROR AL CREAR LA BASE DE DATOS: {e}")
      finally:
        conn.close_connection()



def colums_contabilizado():
    columns = {
    'NOMBRE':'TEXT',
    'CEDULA':'TEXT',
    'PRESTAMO':'TEXT',
    'CUOTA':'CURRENCY',
    'IMPORTE':'CURRENCY',
    'ALTURA_CUOTA':'TEXT',
    'CUOTAS_PAGADAS':'TEXT',
    'PLAZO_RESTANTE':'TEXT',
    'CONVENIO':'TEXT',
    'CONSECUTIVO':'TEXT',
    'OFICINA': 'TEXT',
    'SALDO_CAPITAL':'CURRENCY',
    'TOTAL_ADEUDADO' #28
    'PLAZO_INICIAL':'TEXT',
    'NOMBRE_CONVENIO':'TEXT',
    'SEGUNDA_LIB':'TEXT'
    }

    return columns

def colums_parametria():
    columns = {
    'ALIAS': 'TEXT',
    'DESCRICCION': 'TEXT',
    'NIT': 'INTEGER',
    'CUENTA_ESPERA':'INTEGER',
    'DIA_VENCIMIENTO':'INTEGER',
    'FECHA':'TEXT'
    }

    return columns

def colums_prestamo():
    columns = {
    'PRESTAMO': 'TEXT',
    'FECHA_ULTIMO_VENCIMIENTO': 'TEXT'
    }

    return columns

def colums_residuale():
    columns = {
    'PRESTAMO_10': 'TEXT',
    'PRESTAMO_14': 'TEXT',
    }

    return columns

def colums_sobrante():
    columns = {
    'CONVENIO': 'TEXT',
    'CONSECUTIVO': 'TEXT',
    'NOMBRE_CONVENIO':'TEXT',
    'CEDULA':'TEXT',
    'CODIGO_IDENTIFICACION':'TEXT',
    'CODIGO_VERIFICACION':'TEXT',
    'NOMBRE_CLIENTE':'TEXT',
    'PRESTAMO':'TEXT',
    'OFICINA_OBL':'TEXT',
    'OBLIGACION_14':'TEXT',
    'ESTADO':'TEXT',
    'SEGUNDA_LUI':'TEXT',
    'DIGITO_1':'TEXT',
    'DIGITO_2':'TEXT',
    'LLAVE':'TEXT',
    }

    return columns

def colms_insolvencia():
    columns = {
    'CEDULA': 'TEXT',
    'OBLIGACION': 'TEXT',
    'CONVENIO':'TEXT',
    'TIPO':'TEXT',
    }

    return columns

def colms_cedula_cargue():
    columns = {
    'CEDULA': 'TEXT',
    'PRESTAMO_14': 'TEXT',
    'VALOR':'TEXT',
    'CONTAR':'TEXT',
    'CRUCE':'TEXT',
    }

    return columns

    
        
def delete_bases():
    tablas = ["CONTABILIZADOS", "PARAMETRIA", "PRESTAMO", "RESIDUALES", "SOBRANTES", "INSOLVENCIAS"]

    conn = db.AccessConnection()
    conn.connect()

    for i in tablas:
        conn.delete_table(i)
