#for sql
import pymysql

#database connection
connection = pymysql.connect(host="elholm.eu.mysql",user="elholm_eu",passwd="oM5drxLr",database="elholm_eu" )
cursor = connection.cursor()
# some other statements  with the help of cursor
connection.close()