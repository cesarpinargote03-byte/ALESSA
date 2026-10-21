# Outlook Downloader & API ERP

Herramienta automatizada para **Mesa de Ayuda (Helpdesk)** diseñada para extraer, procesar, filtrar y reportar correos de **Outlook Web (Office 365)** directamente a través de la API interna de Exchange.

Permite gestionar semanalmente las solicitudes de mejoras para el ERP (**4 Express + 1 Con Tiempo**) sin depender de navegación manual ni scrolls en la interfaz web.

---

## 1. Arquitectura del Sistema

```
[ Outlook Office 365 ]
         │  (API Exchange: service.svc?action=FindConversation)
         ▼
[ api_findconversation.py ] ── (sesion.json / headless Playwright)
         │
         ├──► [ conversaciones.json ] (Base histórica incremental)
         │
         ├──► [ exportar_excel.py ] ──► [ Reportes / Correos.xlsx ]
         │
         ▼
[ servidor_api.py (FastAPI) ] (Puerto 8000)
         │
         ▼
[ Swagger UI: http://192.168.6.129:8000/docs ]
```

### Componentes Clave:
* **`api_findconversation.py`**: Motor de ingesta y filtrado multicriterio (Año, Categoría, Remitente).
* **`servidor_api.py`**: Servidor REST con FastAPI para interactuar desde el navegador con Swagger UI.
* **`exportar_excel.py`**: Generador de reportes Excel estilizados con tablas nativas, fechas formateadas y texto del requerimiento (`Preview`).
* **`login.py`**: Asistente visual para renovar credenciales en `sesion.json` cuando expiren las cookies de Microsoft.
* **`Dockerfile` & `docker-compose.yml`**: Empaquetado completo para desplegar como servicio 24/7 en Linux.

---

## 2. Paso a Paso del Despliegue Real

### Configuración del Entorno:
* **Máquina Física (Host):** Windows 10 (con Visual Studio Code y navegador web).
* **Máquina Virtual (VM):** VMware con **Ubuntu Server 26.04.1 LTS**.
* **IP del Servidor:** `192.168.6.129`
* **Usuario:** `cesar`

---

### Paso 1: Instalación de Docker en Ubuntu Server

Dentro de la terminal de Ubuntu Server, se ejecutó el instalador oficial de Docker y se configuraron los permisos de usuario:

```bash
# 1. Descargar e instalar Docker Engine + Docker Compose
curl -fsSL https://get.docker.com -o get-docker.sh && sudo sh get-docker.sh

# 2. Asignar permisos al usuario cesar para ejecutar Docker sin 'sudo'
sudo usermod -aG docker $USER

# 3. Aplicar los cambios de grupo (cerrar sesión y volver a entrar o ejecutar su)
exit
# (Vuelves a iniciar sesión con tu usuario y contraseña)

# 4. Verificar que Docker funciona sin permisos de superusuario
docker ps
```

---

### Paso 2: Transferencia del Proyecto desde Windows 10 a la VM

Desde la terminal de **PowerShell en la máquina física (Windows 10)**, se transfirió la carpeta del proyecto a la máquina virtual mediante SSH/SCP:

```powershell
scp -r D:\cesar_2026\CESAR_2026\Bots\OutlookDownloader cesar@192.168.6.129:~/OutlookDownloader
```

*(Esto copia los archivos de código, `sesion.json`, y la base de datos `conversaciones.json` en segundos).*

---

### Paso 3: Construcción y Arranque del Contenedor en la VM

En la terminal de la máquina virtual (Ubuntu Server):

```bash
# 1. Entrar a la carpeta del proyecto transferida
cd ~/OutlookDownloader

# 2. Construir la imagen e iniciar el servicio en segundo plano
docker compose up -d --build
```

**Resultado esperado:**
```text
[+] up 3/3
✔ Image outlookdownloader-outlook-api Built
✔ Network outlookdownloader_default Created
✔ Container outlook-downloader Started
```

---

### Paso 4: Uso Diario desde la Máquina Física (Windows 10)

En tu computadora física de Windows 10, abre cualquier navegador (Chrome, Edge) e ingresa a:

👉 **`http://192.168.6.129:8000/docs`**

#### Operaciones Disponibles:
1. **`GET /estado`**:
   * Consulta el estado de la sesión y la cantidad total de correos almacenados.
2. **`GET /descargar-excel`**:
   * Permite filtrar por:
     * `categoria`: `SOLICITUD` (o `PENDIENTE`, `CULMINADO`, etc.)
     * `anio`: `2026` (o cualquier año específico)
     * `remitente`: Nombre o correo del solicitante (ej. `@alessa.com.ec`)
   * Al presionar **Execute**, el archivo `.xlsx` generado se descarga automáticamente a tu carpeta de Descargas de Windows.
3. **`POST /sincronizar`**:
   * Se conecta a la API de Outlook en segundo plano para consultar correos nuevos e incorporarlos a la base local sin duplicados.
4. **`GET /consultar`**:
   * Muestra la previsualización de los correos filtrados en formato JSON.

Si `/sincronizar` devuelve `HTTP 401`, vuelve a generar la sesión ejecutando `python login.py` y autentícate hasta ver la bandeja. En Docker, copia el nuevo `sesion.json` al servidor y reinicia el contenedor antes de volver a sincronizar.

---

## 3. Mantenimiento y Comandos Útiles

### Monitoreo en la VM:
* **Ver logs en tiempo real:**
  ```bash
  cd ~/OutlookDownloader
  docker compose logs -f
  ```
* **Reiniciar el contenedor:**
  ```bash
  docker compose restart
  ```
* **Detener el contenedor:**
  ```bash
  docker compose down
  ```
* **Reconstruir tras cambios en el código:**
  ```bash
  docker compose up -d --build
  ```

---

### ¿Qué hacer si expira la sesión de Microsoft 365?

Las cookies de sesión en `sesion.json` son duraderas, pero si la empresa fuerza un nuevo inicio de sesión:

1. En tu máquina con interfaz gráfica (o localmente con Python) ejecutas:
   ```powershell
   python login.py
   ```
2. Inicias sesión en Outlook con tus credenciales y apruebas el Authenticator.
3. Al ver tu bandeja, presionas **ENTER** en la consola para actualizar `sesion.json`.
4. Copias el nuevo `sesion.json` a la VM mediante SCP:
   ```powershell
   scp D:\cesar_2026\CESAR_2026\Bots\OutlookDownloader\sesion.json cesar@192.168.6.129:~/OutlookDownloader/sesion.json
   ```
5. Reinicias el contenedor en la VM:
   ```bash
   docker compose restart
   ```
