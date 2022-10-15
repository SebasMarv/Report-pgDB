"""LIBRERIAS"""
import psycopg2
import xlwt
import time
import smtplib
import glob, os
from datetime import date,timedelta
from email.message import EmailMessage

"""--------------------Super clase (REPORTE)--------------------"""
class REPORTE:

    def __init__(self,usuario,contraseña):
        self.usuario=usuario
        self.contraseña=contraseña

    def Connection_open(self):
        print(f"\n=> PostgreSQL connection is open ✔\n")
        self.connection = psycopg2.connect(
            user=self.usuario,
            password=self.
            contraseña,
            host ="example.com",
            port ="1010",
            database ="postgres"
            )
        self.cur = self.connection.cursor()
        return self.cur

    def Connection_close(self):
        self.connection.close()
        self.cur.close()
        print("\n=> PostgreSQL connection is closed ✔")
    
    def Excel_sheet(self):
        self.wb = xlwt.Workbook()
        self.sheet = self.wb.add_sheet('Reporte')
        return self.sheet
    
    def Excel_save(self,nom_report):
        self.nom = f"{nom_report}.xls"
        print("=> Operation done successfully ✔")
        return self.wb.save(self.nom)

    def Date_start(self):
        self.today_date = date.today()
        td_before=timedelta(1)
        self.fech_start=self.today_date-td_before
        return self.fech_start


    def Date_finish(self):
        self.today_date = date.today()
        td_after=timedelta(0)
        self.fech_finish=self.today_date-td_after
        return self.fech_finish

    def Email_estructura(self):
        self.remitente = "prueba_remitente@hotmail.com"
        self.destinatario = "prueba_destinatario@hotmail.com"
        self.mensaje = f"<p>Estimado,</p><p>Buenos días, se adjunta el reporte solicitado del {self.fech_start} al {self.fech_finish}.</p><p>Saludos,</p>"
        self.email = EmailMessage()
        self.email["From"] = self.remitente
        self.email["To"] = self.destinatario
        self.email["Subject"] = f"Reporte Cliente {self.today_date}"
        self.email.set_content(self.mensaje, subtype="html")

    def Email_archivo(self,nom):
        with open(f"{nom}", "rb") as f:
            self.email.add_attachment(
            f.read(),
            filename=f"{nom}",
            maintype="application",
            subtype="vnd.ms-excel"
            )

    def Email_send(self):
        """smtp-mail.outlook.com => Exclusivo para servicio de correo OUTLOOK"""
        smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
        smtp.starttls()
        smtp.login(self.remitente, "password")
        smtp.sendmail(self.remitente, self.destinatario, self.email.as_string())
        print("\n=> E-mail send ✔\n")
        smtp.quit()

    """Funcion para borrar archivos generados"""
    def Erase_xls(self):
        print("Limpieza de archivos*****")
        for archivo in glob.glob('**/*.xls', recursive=True):
            os.unlink(archivo)
            print("Borrando archivo: ", archivo)

"""--------------------CLases de REPORTE--------------------"""

"""==>REPORTE OC (ORDEN DE COMPRA)"""
class REPORTE_OC(REPORTE):
    def Execution(self):
        self.xls='REPORTE_OC'
        query=f"select * from cliente in ( '99999999999') and fecha between '{REPORTE.Date_start(self)} 08:00:01' and '{REPORTE.Date_finish(self)} 08:00:00' order by fecha desc"
        REPORTE.Connection_open(self)
        REPORTE.Excel_sheet(self)
        self.cur.execute(query)
        rows = self.cur.fetchall()
        """
        Guia: write(row,col, data, style)
        """
        self.sheet.write(0,0,"columna_1")
        self.sheet.write(0,1,"columna_2")
        self.sheet.write(0,2,"columna_3")
        self.sheet.write(0,3,"columna_4")
        self.sheet.write(0,4,"columna_5")
        self.sheet.write(0,5,"columna_6")
        self.sheet.write(0,6,"columna_7")
        self.sheet.write(0,7,"columna_8")
        """Estilo de la fecha""" 
        style_fech=xlwt.easyxf('',num_format_str='DD-MM-YY HH:MM')
        i=1
        for row in rows:
            """
            Guia: write(row,col, data, style)
            """
            self.sheet.write(i,0,row[0])
            self.sheet.write(i,1,row[1])
            self.sheet.write(i,2,row[2])
            self.sheet.write(i,3,row[3])
            self.sheet.write(i,4,row[4])
            self.sheet.write(i,5,row[5])
            self.sheet.write(i,6,row[6])
            self.sheet.write(i,7,row[7],style_fech)
            i=i+1
        REPORTE.Excel_save(self,self.xls)

class REPORTE_GUIA(REPORTE):
    def Execution(self):
        self.xls='REPORTE_GUIAS'
        query=f"select * from cliente in ( '99999999999') and fecha between '{REPORTE.Date_start(self)} 08:00:01' and '{REPORTE.Date_finish(self)} 08:00:00' order by fecha desc"
        REPORTE.Connection_open(self)
        REPORTE.Excel_sheet(self)
        self.cur.execute(query)
        rows = self.cur.fetchall()
        self.sheet.write(0,0,"columna_1")
        self.sheet.write(0,1,"columna_2")
        self.sheet.write(0,2,"columna_3")
        self.sheet.write(0,3,"columna_4")
        self.sheet.write(0,4,"columna_5")
        self.sheet.write(0,5,"columna_6")
        self.sheet.write(0,6,"columna_7")
        self.sheet.write(0,7,"columna_8")
        self.sheet.write(0,7,"columna_8")
        """Estilo de la fecha""" 
        style_fech=xlwt.easyxf('',num_format_str='DD-MM-YY HH:MM')
        i=1
        for row in rows:
            """
            Guia: write(row,col, data, style)
            """
            self.sheet.write(i,0,row[0])
            self.sheet.write(i,1,row[1])
            self.sheet.write(i,2,row[2])
            self.sheet.write(i,3,row[3])
            self.sheet.write(i,4,row[4])
            self.sheet.write(i,5,row[5])
            self.sheet.write(i,6,row[6])
            self.sheet.write(i,7,row[7])
            self.sheet.write(i,8,row[8],style_fech)
            i=i+1
        REPORTE.Excel_save(self,self.xls)

"""Credenciales de conexion a la DB postgres"""
usuario = "user"
constraseña = "password"

"""Inicializar clases"""
general=REPORTE(usuario,constraseña)
reporte_oc=REPORTE_OC(usuario,constraseña)
reporte_guia=REPORTE_GUIA(usuario,constraseña)


"""Fechas necesarias para estructura"""
general.Date_start()
general.Date_finish()

"""Creacion de reporte prueba"""
reporte_oc.Execution()
time.sleep(2)
reporte_guia.Execution()
time.sleep(2)

"""Lista de variables"""
lista_xls=[reporte_oc.nom,reporte_guia.nom]

"""Verificacion de archivos"""
print("\nLista de archivos: \n",lista_xls)

"""Creacion de estructura del e-mail"""
general.Email_estructura()
time.sleep(2)

"""
-Incluir Adjuntos en el envio
-Bucle para tomar los adjuntos
"""
for n in lista_xls:
    general.Email_archivo(n)
time.sleep(2)

"""
Conexion y envio de correo
"""
general.Email_send()
time.sleep(2)

"""
Limpieza archivos con extension .xls en 
la carpeta donde se ejecute la aplicacion
"""
general.Erase_xls()

"""
Cerrar las conexiones abiertas para consultas
"""
reporte_oc.Connection_close()
time.sleep(2)
reporte_guia.Connection_close()
time.sleep(2)