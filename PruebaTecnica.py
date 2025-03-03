from openpyxl import *
import smtplib
import dateparser

class ArchivoXlsx:

    # constructor: inicializamos las variables necesarias
    def __init__(self, file):
        self.file = file
        self.datosEntregas = {}
        self.conteoCiudadesPendientes = {}
        self.entregasProcesadas = 0
        self.montoEntregado = 0

    def leerArchivo(self):
        try:
            # Leer archivo
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
        except FileNotFoundError:
            print("No se encontro el archivo entregas_pendientes.xlsx")

    def guardarArchivo(self):
        # Guardar archivo
        libro = Workbook()
        hoja = libro.active
        # Inicializamos la primera fila del archivo
        hoja['A1'] = "id_entrega"
        hoja['B1'] = "fecha_pedido"
        hoja['C1'] = "cliente"
        hoja['D1'] = "correo_cliente"
        hoja['E1'] = "ciudad"
        hoja['F1'] = "estado_entrega"
        hoja['G1'] = "valor"
        columna = 'A'
        fila = 2
        # Modificamos y guardamos el nuevo archivo
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
        # Recorrido de todas las entregas procesadas para hacer las respectivas transformaciones
        for llave in self.datosEntregas.keys():
            self.datosEntregas[llave][0] = self.transformacionFecha(str(self.datosEntregas[llave][0]))
            self.datosEntregas[llave][1] = self.transformacionCliente(self.datosEntregas[llave][1])
            self.datosEntregas[llave][5] = self.transformacionValor(float(self.datosEntregas[llave][5]))

    def transformacionFecha(self, fecha_str):
        # Cambio de formato de fecha
        fecha = dateparser.parse(fecha_str, languages=['es', 'en'])
        fechaTransformada = fecha.strftime("%Y-%m-%d")
        return fechaTransformada

    def transformacionCliente(self, cliente):
        # Borrar espacios al inicio y al final
        return cliente.strip()

    def transformacionValor(self, valor):
        # Redondear a 2 cifras decimal
        return round(valor, 2)

    def redactarEmails(self):
        for llave in self.datosEntregas.keys():
            # Redactar email entregado
            if self.datosEntregas[llave][4].lower() == "entregado":
                asunto = "Tu pedido ha sido entregado"
                cuerpo = f"Hola {self.datosEntregas[llave][1].strip()},\nTu pedido con ID {llave} ha sido entregado con exito. Gracias por confiar en nosotros."
                mensaje = f"Subject: {asunto}\n\n{cuerpo}"
                self.enviarEmail(mensaje, self.datosEntregas[llave][2])

            # Redactar email pendiente
            elif self.datosEntregas[llave][4].lower() == "pendiente":
                asunto = "Tu pedido esta en camino"
                cuerpo = f"Hola {self.datosEntregas[llave][1].strip()},\nTu pedido con ID {llave} esta en camino y sera entregado pronto."
                mensaje = f"Subject: {asunto}\n\n{cuerpo}"
                self.enviarEmail(mensaje, self.datosEntregas[llave][2])

    def enviarEmail(self, email, destinatario):
        # Conexion con Gmail
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        email_address = 'pruebatecnicaauxprog@gmail.com'
        email_password = 'ozgw ynet vuup xeuu'
        conexion = smtplib.SMTP(smtp_server, smtp_port)
        conexion.ehlo()

        conexion.starttls()
        conexion.login(email_address,email_password)

        # Enviar email
        print(destinatario)
        conexion.sendmail(email_address, destinatario, email)

    def contarCiudades(self):
        # contar ciudades pendientes y calcular el monto total de entregas finalizadas
        for llave in self.datosEntregas.keys():
            if self.datosEntregas[llave][4].lower() == "pendiente":
                if self.conteoCiudadesPendientes.get(self.datosEntregas[llave][3]):
                    self.conteoCiudadesPendientes[self.datosEntregas[llave][3]] += 1
                else:
                    self.conteoCiudadesPendientes[self.datosEntregas[llave][3]] = 1

            if self.datosEntregas[llave][4].lower() == "entregado":
                self.montoEntregado += self.datosEntregas[llave][5]


    def generarReporte(self):
        try:
            reporte = "\n\n------------------------------------------------------------\nResumen\n------------------------------------------------------------\n"
            self.contarCiudades()
            ciudades, pendientes = self.ciudadesPendientes()
            reporte += f"Total de entregas procesadas {self.entregasProcesadas}\nCiudades con mas entregas pendientes:\n"
            for ciudad in ciudades:
                reporte+=f"- {ciudad}: {pendientes}\n"
            reporte += f"Monto total de entregas realizadas: {self.montoEntregado}\n\n"
            print(reporte)
        except ValueError:
            print("Ocurrio un error, no se pudo generar el resumen")

    def ciudadesPendientes(self):
        # Contamos el numero de ciudades pendientes
        maxPendientes = max(self.conteoCiudadesPendientes.values())
        ciudadesMax = [ciudad for ciudad in self.conteoCiudadesPendientes if self.conteoCiudadesPendientes[ciudad] == maxPendientes]
        return ciudadesMax, maxPendientes


xlsx = ArchivoXlsx("Archivos_xlsx/entregas_pendientes")
xlsx.leerArchivo()
xlsx.transformaciones()
xlsx.guardarArchivo()
xlsx.redactarEmails()
xlsx.generarReporte()