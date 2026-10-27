# Guía: acceso directo de escritorio y versión para conductores (celular)

Esta guía tiene 5 partes. Hazlas en orden; cada una se puede probar antes de pasar a la siguiente.

| Parte | Qué logra | Dónde se hace |
|---|---|---|
| 1 | Abrir la app con doble clic, sin CMD | Tu PC |
| 2 | Que el registro de consultas quede en una hoja de Google | Google + tu PC |
| 3 | Subir el proyecto a GitHub (privado) | GitHub |
| 4 | Publicar la versión de conductores en internet | Streamlit Community Cloud |
| 5 | Dejar el ícono en el celular de los conductores | Celular |

---

## Parte 1 — Acceso directo en el escritorio (Windows)

1. Copia todos los archivos de la actualización **dentro de la carpeta de la app** (la que tiene `app.py` y `venv`).
   Acepta reemplazar `app.py` y `utils_pesv.py`. No toques la carpeta `data`.
2. Haz doble clic en **`Crear_acceso_directo.vbs`** (una sola vez). Aparece "App PESV" en tu Escritorio.
3. Doble clic en **App PESV**. La primera vez tarda unos 10 a 20 segundos y luego se abre el navegador solo.

- Para cerrar la app por completo: doble clic en `Detener_PESV.bat`.
- Si no abre: mira el archivo `pesv_log.txt` (en la carpeta de la app); ahí queda el motivo.
- Si mueves la carpeta de la app a otro lugar, vuelve a ejecutar `Crear_acceso_directo.vbs`.
- La forma anterior (CMD) sigue funcionando igual.

## Parte 2 — Registro de consultas en Google Sheets

1. En tu cuenta de Google (mejor la institucional de DELTA GLOBAL) crea una **hoja de cálculo nueva** y ponle un nombre, por ejemplo "Registro consultas PESV".
2. Menú **Extensiones > Apps Script**. Borra el código que aparece y pega todo el contenido de `google_apps_script/Codigo.gs`.
3. En la primera línea de código cambia `CAMBIA_ESTE_TOKEN_POR_UNO_LARGO_Y_PROPIO` por una clave larga inventada por ti
   (mínimo 24 letras y números, sin espacios). **Guárdala**: la necesitas en el paso 6. Guarda el proyecto (ícono del disquete).
4. Botón **Implementar > Nueva implementación**. En el engranaje elige **Aplicación web** y configura:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier persona**
   Pulsa **Implementar**. Google pide autorización: **Autorizar acceso**, elige tu cuenta y, si aparece "Google no ha verificado esta aplicación",
   pulsa **Configuración avanzada > Ir a (nombre del proyecto) > Permitir**. Es normal: la aplicación es tuya.
5. Copia la **URL de la aplicación web** (termina en `/exec`). Pégala en el navegador: debe mostrar `Servicio de registro PESV activo`.
6. En tu PC, dentro de la carpeta `.streamlit` copia `secrets.toml.example` como `secrets.toml` y completa `SHEETS_URL` (paso 5) y `SHEETS_TOKEN` (paso 3).
7. Reinicia la app, entra como Consultar ruta, haz una consulta de prueba y revisa que aparezca una fila nueva en la hoja (pestaña "Consultas").

Si algún día cambias el código del script: **Implementar > Administrar implementaciones > lápiz > Versión: Nueva versión > Implementar**.
Así la URL no cambia. Si en cambio creas una "Nueva implementación", la URL sí cambia y hay que actualizarla en los secretos.

## Parte 3 — Subir el proyecto a GitHub

1. En github.com crea un repositorio **Private** (privado), por ejemplo `pesv-rutas`.
2. Instala **GitHub Desktop**, inicia sesión y agrega la carpeta de la app (File > Add local repository; si pide, "create a repository").
   GitHub Desktop respeta el archivo `.gitignore`, por eso no sube `venv`, ni `secrets.toml`, ni el registro de consultas con cédulas.
3. Antes de publicar revisa en la lista de archivos que **no** aparezcan `venv`, `secrets.toml`, `consultas.json` ni `pesv_log.txt`.
4. **Commit** y **Publish repository** (deja marcada la casilla "Keep this code private").
5. `config.py` contiene la contraseña de administrador. En la versión publicada no se usa, pero cámbiala por una que no uses en otros sitios.

## Parte 4 — Publicar la versión de conductores (Streamlit Community Cloud)

1. Entra a share.streamlit.io con tu cuenta de GitHub y autoriza el acceso a tus repositorios (incluidos los privados).
2. **Create app** > repositorio `pesv-rutas`, rama `main`, archivo principal `app.py`.
3. **Advanced settings**: elige Python **3.12** y en **Secrets** pega:

```
MODO_PUBLICO = true
SHEETS_URL = "la URL que termina en /exec"
SHEETS_TOKEN = "tu token"
```

4. **Deploy**. Cuando termine, la app queda en una dirección tipo `https://pesv-rutas.streamlit.app/`. Ábrela y haz una consulta de prueba.
   En esta versión **no existe** el acceso de administrador: solo se ve "Consultar ruta".
5. Las apps gratuitas se "duermen" tras un tiempo sin visitas; al abrirlas piden pulsar un botón y tardan unos segundos en despertar.

## Parte 5 — Ícono en el celular de los conductores

- **Android (Chrome):** abrir la dirección > menú de tres puntos > "Agregar a la pantalla principal" (o "Instalar app").
- **iPhone (Safari):** abrir la dirección > botón Compartir > "Añadir a pantalla de inicio".

Con Streamlit esto crea un acceso directo que abre la app; no es una aplicación instalada de la tienda.
El nombre y el ícono pueden variar según el teléfono.

---

## Cómo actualizar las rutas después de publicar

1. En tu PC (App PESV de escritorio, como administrador) crea o edita las rutas y sube las capturas del mapa.
2. En GitHub Desktop aparecerán cambios en `data/rutas.json` y `data/mapas/`. Escribe un resumen y pulsa **Commit** y luego **Push origin**.
3. En uno o dos minutos la versión de los conductores se actualiza sola.

## Dónde ver el registro de consultas

- Directamente en la hoja de Google (puedes filtrar, descargar o compartirla con auditoría).
- En la app de escritorio, menú **Registro de consultas** (lee la misma hoja si tienes `secrets.toml` configurado).

## Datos personales

El registro guarda nombre y cédula de los conductores. Confirma con el área jurídica de la empresa el aviso de tratamiento de datos
(Ley 1581 de 2012) y ajusta el texto que ven los conductores con el ajuste `AVISO_DATOS` (ver `ajustes.py`).
Comparte la hoja de Google solo con quienes deban consultarla.
