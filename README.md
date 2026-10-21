# 🚀 NeXo Automation - Sistema de Prospección Comercial & Emailing

<div align="center">
  <img src="assets/logo.png" alt="NeXo Software Logo" width="160" />
  <br><br>
  
  [![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
  [![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-blue?style=for-the-badge)](https://www.python.org/)
  [![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](#)
  [![Status](https://img.shields.io/badge/Status-Active-brightgreen?style=for-the-badge)](#)
  
  <p align="center">
    <strong>Herramienta integral de automatización para prospección comercial, gestión de leads en Excel y envío segmentado de correos electrónicos vía Gmail SMTP.</strong>
  </p>
</div>

---

## 📋 Tabla de Contenidos

- [Descripción General](#-descripción-general)
- [Características Principales](#-características-principales)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Requisitos Previos](#-requisitos-previos)
- [Instalación y Configuración](#-instalación-y-configuración)
  - [1. Clonar el Repositorio](#1-clonar-el-repositorio)
  - [2. Crear un Entorno Virtual](#2-crear-un-entorno-virtual)
  - [3. Instalar Dependencias](#3-instalar-dependencias)
  - [4. Configurar Variables de Entorno (.env)](#4-configurar-variables-de-entorno-env)
- [Configuración de Gmail (Contraseña de Aplicación)](#-configuración-de-gmail-contraseña-de-aplicación)
- [Estructura del Archivo Excel](#-estructura-del-archivo-excel)
- [Modo de Uso](#-modo-de-uso)
  - [Modo 1: Interfaz Gráfica (GUI) - Recomendado](#modo-1-interfaz-gráfica-gui---recomendado)
  - [Modo 2: Modo Consola / Terminal](#modo-2-modo-consola--terminal)
- [Buenas Prácticas y Entregabilidad](#-buenas-prácticas-y-entregabilidad)
- [Seguridad](#-seguridad)
- [Contribuciones y Licencia](#-contribuciones-y-licencia)

---

## 💡 Descripción General

**NeXo Automation** es una solución diseñada para equipos de ventas, agencias y profesionales independientes que gestionan prospectos comerciales en hojas de cálculo y necesitan ejecutar campañas de contacto en frío (cold emailing) de manera profesional, automatizada y personalizada.

El sistema lee los contactos de un libro de Excel (`NeXo_Gestion_Prospectos_y_Correos.xlsx`), filtra por categorías o sectores, valida la estructura de los correos, envía mensajes formateados en texto plano y HTML usando un servidor SMTP seguro de Gmail, y actualiza el estado de la gestión directamente en el archivo Excel con marca de tiempo y trazabilidad completa.

---

## ✨ Características Principales

- 🖥️ **Interfaz Gráfica Moderna (Dark Theme):** Desarrollada con Tkinter y ttk, inspirada en una estética Slate/Sky elegante y funcional con métricas en tiempo real.
- 🎯 **Segmentación de Prospectos:** Permite filtrar y disparar campañas por categorías específicas (e.g. *Hoteles y hospedajes*, *Restaurantes*, etc.) o enviar a todos los pendientes.
- ⏱️ **Retrasos Dinámicos Anti-Spam:** Rango de pausa configurable y aleatorio entre envíos (por defecto 45–75 segundos) para simular comportamiento humano y proteger la reputación del remitente.
- 📬 **Validación y Limpieza de Destinatarios:** Detección de correos múltiples separados por comas o punto y coma, filtrado de direcciones inválidas.
- 📝 **Actualización Automática de Excel:** Escribe en la hoja de cálculo el estado del contacto (`Enviado`, `Error`), fecha/hora exacta y detalle del resultado.
- 🔒 **Modo Simulación / Prueba de Conexión:** Botón de verificación SMTP instantánea y previsualización de destinatarios antes de lanzar la campaña.
- 🛑 **Parada Segura en Caliente:** Posibilidad de pausar o detener la campaña en cualquier momento sin corromper el archivo Excel.
- 💻 **Doble Modo de Ejecución:** Vía interfaz de escritorio moderna o mediante script de consola desatendido.

---

## 📂 Estructura del Proyecto

```plaintext
nexo-automation/
│
├── assets/
│   └── logo.png                               # Logotipo de NeXo Software
│
├── .env.example                               # Plantilla de variables de entorno
├── .gitignore                                 # Exclusiones de Git (protege credenciales)
├── app_gui.py                                 # Interfaz gráfica de usuario (Tkinter Dark Mode)
├── config.py                                  # Carga y validación de variables de configuración
├── enviar_campana.py                          # Lógica central del motor de envío y manejo de Excel
├── ejecutar.bat                               # Lanzador rápido con 1-click para Windows (GUI)
├── ejecutar_consola.bat                       # Lanzador por consola para Windows
├── NeXo_Gestion_Prospectos_y_Correos.xlsx     # Plantilla de base de datos de prospectos
├── requirements.txt                           # Dependencias del proyecto
└── README.md                                  # Documentación del repositorio
```

---

## 🛠️ Requisitos Previos

- **Python:** Versión `3.10` o superior instalada ([python.org](https://www.python.org/downloads/)).
- **Sistema Operativo:** Windows 10/11, macOS o distribuciones Linux compatibles.
- **Cuenta de Gmail:** Con autenticación en dos pasos activa para generar una contraseña de aplicación.

---

## 🚀 Instalación y Configuración

### 1. Clonar el Repositorio

```bash
git clone https://github.com/Novato777/Automatizaci-n-.git
cd Automatizaci-n-
```

### 2. Crear un Entorno Virtual

Es recomendable utilizar un entorno virtual para aislar las librerías:

```bash
# En Windows:
python -m venv venv
venv\Scripts\activate

# En macOS / Linux:
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar Dependencias

Instala los paquetes necesarios utilizando `requirements.txt`:

```bash
pip install -r requirements.txt
```

*(Nota: La librería `openpyxl` es la única dependencia externa; Tkinter y las librerías de red/email están incluidas en la biblioteca estándar de Python).*

### 4. Configurar Variables de Entorno (`.env`)

Copia el archivo de ejemplo para crear tu archivo `.env` privado:

```bash
# En Windows (CMD / PowerShell):
copy .env.example .env

# En Linux / macOS:
cp .env.example .env
```

Abre `.env` con un editor de texto y completa los valores correspondientes:

```env
GMAIL_USER=tu_correo@gmail.com
GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
SENDER_NAME=NeXo Software
EXCEL_PATH=NeXo_Gestion_Prospectos_y_Correos.xlsx
DELAY_MIN_SECONDS=45
DELAY_MAX_SECONDS=75
```

> [!WARNING]
> **NUNCA** subas tu archivo `.env` a GitHub. El archivo `.gitignore` ya está configurado para proteger tus credenciales.

---

## 🔐 Configuración de Gmail (Contraseña de Aplicación)

Google no permite usar tu contraseña habitual para conexiones SMTP externas. Debes generar una **Contraseña de Aplicación de 16 letras**:

1. Ingresa a tu cuenta de Google y ve a [Seguridad de Google](https://myaccount.google.com/security).
2. Asegúrate de tener activada la **Verificación en dos pasos (2FA)**.
3. En la barra de búsqueda de la cuenta de Google, escribe **"Contraseñas de aplicaciones"** (o navega a [esta dirección directa](https://myaccount.google.com/apppasswords)).
4. Asigna un nombre a la aplicación (por ejemplo: `NeXo Automation`).
5. Haz clic en **Crear**. Google te mostrará un código de 16 caracteres (ej: `abcd efgh ijkl mnop`).
6. Copia este código y pégalo en la variable `GMAIL_APP_PASSWORD` dentro de tu archivo `.env`.

---

## 📊 Estructura del Archivo Excel

El archivo `NeXo_Gestion_Prospectos_y_Correos.xlsx` utiliza una hoja llamada **`Contactos`** con la siguiente disposición de columnas:

| Fila / Columna | Columna (Letra) | Nombre del Campo | Propósito |
| :--- | :---: | :--- | :--- |
| Fila 4 | - | Encabezados de tabla | Nombres de columnas |
| Fila 5+ | Columna 1 (A) | `ID` | Identificador único del lead |
| Fila 5+ | Columna 2 (B) | `Negocio / Empresa` | Nombre comercial del destinatario |
| Fila 5+ | Columna 3 (C) | `Categoría` | Segmento (ej: *Hoteles y hospedajes*) |
| Fila 5+ | Columna 4 (D) | `Sector` | Sector económico o ciudad |
| Fila 5+ | Columna 5 (E) | `Correo` | Correo o correos (separados por coma) |
| Fila 5+ | Columna 12 (L)| `Asunto` | Asunto personalizado del correo |
| Fila 5+ | Columna 13 (M)| `Cuerpo` | Contenido del mensaje |
| Fila 5+ | Columna 16 (P)| `Estado Campaña` | Estado actual (`Pendiente`, `Enviado`, etc.) |
| Fila 5+ | Columna 17 (Q)| `Fecha Envío` | Registrada automáticamente tras el envío |
| Fila 5+ | Columna 18 (R)| `Resultado / Detalle`| Log del resultado (`OK`, descripción de error) |

---

## 🖥️ Modo de Uso

### Modo 1: Interfaz Gráfica (GUI) - Recomendado

Puedes iniciar la interfaz de dos maneras:

- **En Windows:** Haz doble clic sobre el archivo [`ejecutar.bat`](ejecutar.bat).
- **Desde la terminal:**
  ```bash
  python app_gui.py
  ```

**Flujo en la Interfaz:**
1. **Comprobar Conexión:** Haz clic en *"Probar Conexión SMTP"* para asegurar que tus credenciales de Gmail son válidas.
2. **Seleccionar Categoría:** Escoge la categoría que deseas contactar o selecciona *"Todas"*.
3. **Definir Límite y Tiempos:** Ajusta el número máximo de correos a enviar y el intervalo de espera en segundos.
4. **Cargar Prospectos:** Presiona *"Cargar Prospectos"* para auditar la lista de destinatarios antes de enviar.
5. **Iniciar Campaña:** Haz clic en *"Iniciar Campaña"*. Observarás el avance en la barra de progreso y el registro en la consola integrada.

---

### Modo 2: Modo Consola / Terminal

Si deseas ejecutar envíos automatizados mediante tareas programadas (Cron en Linux o Task Scheduler en Windows):

- **En Windows:** Haz doble clic sobre [`ejecutar_consola.bat`](ejecutar_consola.bat).
- **Desde la terminal:**
  ```bash
  python enviar_campana.py
  ```

El script te solicitará interactivamente la categoría a enviar, el límite y confirmará la ejecución mostrando el progreso paso a paso.

---

## 📈 Buenas Prácticas y Entregabilidad

Para evitar que tus correos terminen en la carpeta de Spam o que Google restrinja temporalmente tu cuenta:

1. **Respeta los Límites de Gmail:**
   - Cuentas `@gmail.com` gratuitas: Máximo 100–150 correos diarios para prospección segura (límite técnico teórico: 500/día).
   - Cuentas de Google Workspace empresariales: Hasta 2,000 correos diarios.
2. **Usa Retrasos Adecuados:** Mantén el delay aleatorio entre 45 y 75 segundos. Enviar decenas de correos en pocos segundos disparará los filtros de abuso de Google.
3. **Personaliza Asuntos y Mensajes:** Evita mensajes genéricos idénticos; personaliza el nombre de la empresa y la propuesta de valor.
4. **Limpia tu Base de Datos:** Elimina correos inexistentes o con rebotes constantes para conservar una buena reputación de dominio.

---

## 🛡️ Seguridad

- Las credenciales nunca se guardan en el código fuente. Se leen dinámicamente desde variables de entorno locales mediante [`config.py`](config.py).
- La conexión con Gmail se realiza exclusivamente mediante **TLS/SSL seguro** (`smtplib.SMTP_SSL` en el puerto `465`).
- Todo cambio en el archivo Excel se guarda de forma segura con respaldo ante posibles interrupciones.

---

## 📄 Licencia

Este proyecto se encuentra bajo la Licencia [MIT](LICENSE). Siéntete libre de utilizarlo, modificarlo y distribuirlo para tus proyectos comerciales o personales.

---

<div align="center">
  Desarrollado con ❤️ para <strong>NeXo Software</strong>
</div>
