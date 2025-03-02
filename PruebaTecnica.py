from openpyxl import *
import smtplib
import dateparser

class ArchivoXlsx:

    def __init__(self, file):
        self.file = file
        self.datosEntregas = {}
        self.conteoCiudadesPendientes = {}
        self.entregasProcesadas = 0
        self.montoEntregado = 0

    def leerArchivo(self):
        libro = load_workbook(f"{self.file}.xlsx")
        hoja = libro["Hoja1"]
        columna = 'A'
        fila = 2
        while hoja[f"{columna}{fila}"].value != None:
            while columna != 'H':
                if columna == 'A':
                    self.datosEntregas[hoja[f"{columna}{fila}"].value] = []
                else:
                    self.datosEntregas[hoja[f"A{fila}"].value].append(hoja[f"{columna}{fila}"].value)
                columna = chr(ord(columna) + 1)

            fila += 1
            columna = 'A'

    def guardarArchivo(self):
        libro = Workbook()
        hoja = libro.active
        hoja['A1'] = "id_entrega"
        hoja['B1'] = "fecha_pedido"
        hoja['C1'] = "cliente"
        hoja['D1'] = "correo_cliente"
        hoja['E1'] = "ciudad"
        hoja['F1'] = "estado_entrega"
        hoja['G1'] = "valor"
        columna = 'A'
        fila = 2
        for key in self.datosEntregas.keys():
            if self.datosEntregas[key][4].lower() != "devuelto":
                while columna != 'H':
                    if columna == 'A':
                        hoja[f"{columna}{fila}"].value = key
                        self.entregasProcesadas+=1
                    else:
                        hoja[f"{columna}{fila}"].value = self.datosEntregas[key][ord(columna)-66]
                    columna = chr(ord(columna) + 1)
                fila += 1
                columna = 'A'

        libro.save(f"{self.file}_modificado.xlsx")

    def transformaciones(self):
        for llave in self.datosEntregas.keys():
            self.datosEntregas[llave][0] = self.transformacionFecha(str(self.datosEntregas[llave][0]))
            self.datosEntregas[llave][1] = self.transformacionCliente(self.datosEntregas[llave][1])
            self.datosEntregas[llave][5] = self.transformacionValor(float(self.datosEntregas[llave][5]))

    def transformacionFecha(self, fecha_str):
        fecha = dateparser.parse(fecha_str, languages=['es', 'en'])
        fechaTransformada = fecha.strftime("%Y-%m-%d")
        return fechaTransformada

    def transformacionCliente(self, cliente):
        return cliente.strip()

    def transformacionValor(self, valor):
        return round(valor, 2)

    def redactarEmails(self):
        for llave in self.datosEntregas.keys():
            if self.datosEntregas[llave][4].lower() == "entregado":
                asunto = "Tu pedido ha sido entregado"
                cuerpo = f"Hola {self.datosEntregas[llave][1].strip()},\nTu pedido con ID {llave} ha sido entregado con exito. Gracias por confiar en nosotros."
                mensaje = f"Subject: {asunto}\n\n{cuerpo}"
                self.enviarEmail(mensaje, self.datosEntregas[llave][2])

            elif self.datosEntregas[llave][4].lower() == "pendiente":
                asunto = "Tu pedido esta en camino"
                cuerpo = f"Hola {self.datosEntregas[llave][1].strip()},\nTu pedido con ID {llave} esta en camino y sera entregado pronto."
                mensaje = f"Subject: {asunto}\n\n{cuerpo}"
                self.enviarEmail(mensaje, self.datosEntregas[llave][2])

    def enviarEmail(self, email, destinatario):
        conexion = smtplib.SMTP('smtp.gmail.com', 587)
        conexion.ehlo()

        conexion.starttls()

        conexion.login('pruebatecnicaauxprog@gmail.com','ozgw ynet vuup xeuu')

        print(destinatario)
        conexion.sendmail('pruebatecnicaauxprog@gmail.com', destinatario, email)

    def contarCiudades(self):
        for llave in self.datosEntregas.keys():
            if self.datosEntregas[llave][4].lower() == "pendiente":
                if self.conteoCiudadesPendientes.get(self.datosEntregas[llave][3]):
                    self.conteoCiudadesPendientes[self.datosEntregas[llave][3]] += 1
                else:
                    self.conteoCiudadesPendientes[self.datosEntregas[llave][3]] = 1

            if self.datosEntregas[llave][4].lower() == "entregado":
                self.montoEntregado += self.datosEntregas[llave][5]


    def generarReporte(self):
        reporte = "\n\n------------------------------------------------------------\nResumen\n------------------------------------------------------------\n"
        self.contarCiudades()
        ciudades, pendientes = self.ciudadesPendientes()
        reporte += f"Total de entregas procesadas {self.entregasProcesadas}\nCiudades con mas entregas pendientes:\n"
        for ciudad in ciudades:
            reporte+=f"- {ciudad}: {pendientes}\n"
        reporte += f"Monto total de entregas realizadas: {self.montoEntregado}\n\n"
        print(reporte)

    def ciudadesPendientes(self):
        maxPendientes = max(self.conteoCiudadesPendientes.values())
        ciudadesMax = [ciudad for ciudad in self.conteoCiudadesPendientes if self.conteoCiudadesPendientes[ciudad] == maxPendientes]
        return ciudadesMax, maxPendientes


xlsx = ArchivoXlsx("Archivos_xlsx/entregas_pendientes")
xlsx.leerArchivo()
xlsx.transformaciones()
xlsx.guardarArchivo()
xlsx.redactarEmails()
xlsx.generarReporte()