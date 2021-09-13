#-------------------------------------------------------------------------------
# Name:        Examen
#
# Author:      Arturo Hernández Gómez
# Date:        13/09/2021
#-------------------------------------------------------------------------------

import mysql.connector
from zipfile import ZipFile
import glob
import xlrd


#Conexion a MySQL
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="arturo"
)

#Extraccion de los archivos a la carpeta "Files"
carpeta = "Files.zip"
with ZipFile(carpeta, 'r') as zip:
    zip.extractall("Files")


#Lectura de los archivos
archivos = glob.glob("Files/*")

#lista donde se almacenan las entradas
lista = []

for a in archivos:
    print(a)
    
    #Lectura de los archivos .xls
    f = xlrd.open_workbook(a) 
    hoja = f.sheet_by_index(0)

    #Panel number
    PN = hoja.cell_value(2,3)
    #Job number 
    JN = hoja.cell_value(3,3)
    #Job name 
    JNa = hoja.cell_value(4,3)
    #Seal

    S = hoja.cell_value(2,9)
    if(S=='X'):
        S=1
    else:
        S=0
    #Type
    T = hoja.cell_value(27,1)
    #Modbus
    M = hoja.cell_value(32,2)
    
    #Meter no. 
    Mn = hoja.cell_value(81,7)
    Mn = Mn[10:len(Mn)]

    #Serial number
    SN = hoja.cell_value(49,2)
    i = 50
    while(SN != ""):
        lista.append((SN, PN, JN, JNa, S, T, M, Mn))
        SN = hoja.cell_value(i,2)
        i = i+1
        
#Creacion de la base de datos
mycursor = mydb.cursor()
mycursor.execute("CREATE DATABASE examen")

#Conexion a la base de datos
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="arturo",
  database="examen"
)

mycursor = mydb.cursor()

#creacion de la tabla
mycursor.execute("CREATE TABLE job_traveler (serial_number VARCHAR(255) PRIMARY KEY,\
                                            panel_number VARCHAR(255),\
                                            job_number VARCHAR(255),\
                                            job_name VARCHAR(255),\
                                            seal BINARY(1),\
                                            type VARCHAR(255),\
                                            modbus_id INT,\
                                            meter_no VARCHAR(255))")

#insercion de los datos
sql = "INSERT INTO job_traveler (serial_number,\
                                panel_number,\
                                job_number,\
                                job_name,\
                                seal,\
                                type,\
                                modbus_id,\
                                meter_no)\
                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"

mycursor.executemany(sql, lista)
mydb.commit()

#consulta de los datos insertados
mycursor.execute("SELECT * FROM job_traveler")

myresult = mycursor.fetchall()

for x in myresult:
    print(x)