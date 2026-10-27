# App PESV — Rutas Frecuentes y Riesgos Viales

Aplicación local (Python + Streamlit) para gestionar de forma dinámica las
fichas de rutas frecuentes del Plan Estratégico de Seguridad Vial: datos
generales, riesgos, medidas preventivas, puntos críticos por nivel de
riesgo, y enlace al mapa. Reemplaza el trabajo manual en Excel por
formularios, y permite exportar de nuevo a Excel cuando lo necesites.

---

## 1. Requisitos previos

- **Python 3.9 o superior** instalado en tu computador.
  - Verifícalo abriendo una terminal (CMD/PowerShell en Windows, Terminal en Mac)
    y escribiendo:
    ```
    python --version
    ```
    Si no lo tienes, descárgalo desde https://www.python.org/downloads/
    (en Windows, marca la casilla **"Add Python to PATH"** durante la instalación).

## 2. Descomprimir el proyecto

Descomprime la carpeta `pesv_app` que te entregué en el lugar donde quieras
tenerla, por ejemplo `Documentos\pesv_app`.

## 3. Abrir una terminal dentro de la carpeta

- **Windows:** entra a la carpeta `pesv_app` en el explorador de archivos,
  haz clic en la barra de direcciones, escribe `cmd` y presiona Enter.
- **Mac:** clic derecho sobre la carpeta > "Nueva terminal en la carpeta"
  (o abre Terminal y usa `cd` hasta llegar a la carpeta).

## 4. Crear un entorno virtual (recomendado, evita conflictos)

```bash
python -m venv venv
```

Actívalo:

- **Windows (CMD):**
  ```
  venv\Scripts\activate
  ```
- **Windows (PowerShell):**
  ```
  venv\Scripts\Activate.ps1
  ```
- **Mac / Linux:**
  ```
  source venv/bin/activate
  ```

Verás `(venv)` al inicio de la línea de la terminal cuando esté activo.

## 5. Instalar las dependencias

```bash
pip install -r requirements.txt
```

Esto instala Streamlit, pandas y openpyxl (las únicas librerías que usa la app).

## 6. (Opcional) Migrar tu Excel actual

Si quieres partir de las rutas que ya tienes en
`RUTAS_FRECUENTES_PESV.xlsx`, copia ese archivo dentro de la carpeta
`pesv_app` y ejecuta:

```bash
python migrar_excel.py RUTAS_FRECUENTES_PESV.xlsx
```

Esto crea `data/rutas.json` con tus rutas ya cargadas. Si no lo haces,
la app simplemente empieza vacía y tú vas creando las rutas desde el
formulario.

> Nota: ya te dejo `data/rutas.json` incluido con tus 2 rutas actuales
> (Cali–Arroyohondo y Cali–Medellín) migradas automáticamente, así que
> este paso es opcional a menos que quieras volver a migrar un Excel
> actualizado.

## 7. Ejecutar la aplicación

```bash
streamlit run app.py
```

Esto abre automáticamente tu navegador en `http://localhost:8501`. Si no
se abre solo, copia esa dirección y pégala en tu navegador.

Para **detener la app**, vuelve a la terminal y presiona `Ctrl + C`.

## 8. Usos posteriores (ya instalado)

Cada vez que quieras volver a usar la app, solo necesitas:

```bash
cd ruta/a/pesv_app
venv\Scripts\activate      (Windows)   ó   source venv/bin/activate   (Mac/Linux)
streamlit run app.py
```

---

## ¿Qué puedes hacer en la app?

- **🔎 Consultar ruta:** vista para el conductor — escribe su nombre, elige la
  ruta en un desplegable y pulsa "Consultar ruta". Se muestra la ficha completa
  con el **mapa de Google My Maps embebido**, los puntos críticos, medidas
  preventivas, y un botón para **descargar la ficha en PDF**. Cada consulta
  queda registrada automáticamente (fecha, hora, conductor, ruta) como
  evidencia de que el riesgo fue informado antes del viaje.
- **📋 Listado de rutas:** vista de administrador — ver todas las fichas,
  buscar por nombre/origen/destino, filtrar por nivel de riesgo, editar,
  eliminar o descargar el PDF de cada una.
- **➕ Nueva ruta:** formulario para crear una ficha completa (datos generales,
  riesgos, medidas preventivas y tabla de puntos críticos con nivel Bajo/Medio/Alto).
- **📝 Registro de consultas:** tabla con todas las consultas hechas por los
  conductores (quién, qué ruta, cuándo), con filtros y descarga en CSV —
  es la evidencia de trazabilidad para el PESV.
- **📊 Exportar / Respaldo:** descarga todo en un Excel con el mismo formato de
  fichas (una hoja por ruta, coloreado por nivel de riesgo), o en un JSON de
  respaldo que puedes volver a cargar después.

## Cómo funciona el mapa embebido

Al crear o editar una ruta, pega el link de "compartir" de tu mapa de Google
My Maps (el mismo que ya tenías, tipo `.../d/edit?mid=XXXX...`). La app lo
convierte automáticamente al formato de "insertar" (`.../d/u/0/embed?mid=XXXX`)
para mostrarlo interactivo dentro de la pantalla de consulta. No necesitas
generar tú el código del iframe.

## Estructura del proyecto

```
pesv_app/
├── app.py               → aplicación principal (ejecutar con streamlit run)
├── utils_pesv.py          → mapa embebido, generación de PDF, registro de consultas
├── migrar_excel.py       → script para importar un Excel existente
├── requirements.txt      → dependencias
├── data/
│   ├── rutas.json           → base de datos de rutas (se crea/actualiza sola)
│   └── consultas.json       → registro de consultas de conductores (se crea sola)
└── README.md
```

Todos los datos se guardan localmente en tu propio computador (`data/rutas.json`
y `data/consultas.json`); nada se envía a internet, salvo la carga del mapa de
Google My Maps que ya era pública mediante el link que compartiste.

## Problemas comunes

- **"python no se reconoce como comando":** Python no quedó en el PATH.
  Reinstala marcando "Add Python to PATH", o usa `py` en vez de `python`.
- **El navegador no abre solo:** entra manualmente a `http://localhost:8501`.
- **Quiero cambiar el puerto:** `streamlit run app.py --server.port 8502`.
