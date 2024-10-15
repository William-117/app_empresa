# importar libreria mysql.connector para conectar python con mysql
import mysql.connector


def connectionBD():
    try: 
        #connection = mysql.connector.connect(
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            passwd="William1192",
            database="app_empresa_bd",
            charset='utf8mb4',
            collation='utf8mb4_unicode_ci',
            raise_on_warnings=True
        )
        if connection.is_connected():
            #print("conexion exitosa a la BD")
            return connection
        
    except mysql.connector.Error as error:
        print(f"No se pudo conectar: {error}")