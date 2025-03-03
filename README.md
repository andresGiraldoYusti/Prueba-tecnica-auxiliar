# Prueba-tecnica-auxiliar

Este es un script de Python diseñado para automatizar tareas relacionadas con la lectura, transformación y envío de datos desde un archivo de Excel. El script realiza las siguientes tareas:

1. **Lee un archivo de Excel** para extraer información de una base de datos.
2. **Realiza transformaciones** en los datos extraídos.
3. **Genera un nuevo archivo de Excel** con los datos relevantes y sus transformaciones.
4. **Envía correos electrónicos** de manera automática con las especificaciones proporcionadas.
5. **Muestra un resumen** en la consola con los resultados del proceso.

## Instrucciones para Ejecutar el Script

### Prerrequisitos

Antes de ejecutar el script, asegúrate de tener lo siguiente:

1. **Python 3.x instalado**:
   - Si no lo tienes, descárgalo e instálalo desde [python.org](https://www.python.org/downloads/).
   - Verifica que Python esté instalado correctamente ejecutando en la terminal:
     ```
     python --version
     ```

2. **Librerías necesarias**:
   - El script usa las librerías `openpyxl` y `dateparser`. Instálalas ejecutando:
     ```
     pip install openpyxl dateparser
     ```

---

### 🔧 Configuración

1. **Prepara el archivo de Excel**:
   - Coloca el archivo de Excel (`entregas_pendientes.xlsx`) en la carpeta (`Archivos_xlsx`).
   - Asegúrate de que el archivo tenga la estructura esperada por el script.
     | ID_Entrega | Fecha_Pedido   | Cliente | Correo_Cliente | Ciudad | Estado_Entrega | Valor  |
     |------------|----------------|---------|----------------|--------|---------------|---------|
     |    1001    |  2025/02/15    |   Juan  | juan@email.com |Medellín|   Entregado   |10.000,50|
     |    1002    |  15-02-2025    |   Ana   | ana@email.com  | Bogotá |   Pendiente   | 5.000,00|
     |    1003    |Febrero 15, 2025|  Carlos |carlos@email.com|  Cali  |   Devuelto    |12.500,75|

2. **Asegurate de que esten configuradas las credenciales de correo**:
   - Abre el archivo `PruebaTecnica.py` en un editor de texto o IDE.
   - Busca las siguientes líneas y modifícalas con tus datos:
     ```
     smtp_server = 'smtp.gmail.com'
     smtp_port = 587
     email_address = 'pruebatecnicaauxprog@gmail.com'
     email_password = 'ozgw ynet vuup xeuu'
     ```
   - **Nota:** Si usas Gmail, necesitarás una "contraseña de aplicación" para acceder al servidor SMTP. Puedes generarla en la configuración de tu cuenta de Google.

---

### 🏃‍♂️ Ejecución

Sigue estos pasos para ejecutar el script:

1. **Abre una terminal**:
   - En Windows: Presiona `Win + R`, escribe `cmd` y presiona Enter.
   - En macOS o Linux: Abre la aplicación "Terminal".

2. **Navega a la carpeta del proyecto**:
   - Usa el comando `cd` para moverte a la carpeta donde está el script. Por ejemplo:
     ```
     cd C:\Users\tu_usuario\PruebaTecnica
     ```
     (Reemplaza la ruta con la ubicación real de tu proyecto).

3. **Ejecuta el script**:
   - Una vez en la carpeta correcta, ejecuta el script con el siguiente comando:
     ```
     python PruebaTecnica.py
     ```

4. **Revisa los resultados**:
   - El script hará lo siguiente:
     1. Leerá el archivo `Archivos.xlsx`.
     2. Transformará los datos según las reglas definidas en el script.
     3. Generará un nuevo archivo con los datos transformados (`entregas_pendiente_modificado.xlsx`).
     4. Enviará correos electrónicos usando las credenciales proporcionadas.
     5. Mostrará un resumen del proceso en la consola.