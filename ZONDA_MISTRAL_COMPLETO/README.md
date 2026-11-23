# ZONDA - MISTRAL | Aplicación integrada

Esta entrega integra la aplicación web existente con el motor Python de lectura de comprobantes y generación del Excel basado en la plantilla original `RENDICION DE CUENTAS N°8 OSCAR.xlsx`.

## Funciones incluidas
- Inicio de sesión y roles administrador/usuario.
- Cada usuario ve sus rendiciones; el administrador puede ver todas.
- Creación de rendiciones y carga múltiple de PDF.
- Validación básica de archivos PDF.
- Lectura de texto PDF con PyMuPDF y extracción de fecha, RUC, tipo, serie, número, proveedor, detalle y total.
- Clasificación sugerida en HOSPEDAJE, MOVILIDAD, ALIMENTO u OTROS.
- Datos de rendición configurables desde la web (técnico, DNI, cuenta, importe, fechas, ubicación y descripción).
- Vista de datos reconocidos y corrección manual por comprobante.
- Generación y descarga de Excel usando la plantilla original, sin sobrescribirla.
- Historial de actividad por rendición.
- Persistencia local con SQLite.

## Instalación en Windows
1. Descomprime toda la carpeta en una ubicación permanente.
2. Instala Python 3.11 o superior y activa la opción **Add Python to PATH**.
3. Ejecuta `INICIAR_ZONDA_MISTRAL.bat`. El iniciador instala automáticamente las dependencias, incluida Pillow (necesaria para insertar el logotipo en el Excel).
4. Se abrirá el navegador en `http://127.0.0.1:5000`.
5. Ingresa con `admin` / `Admin123!`.
6. Cambia la contraseña inicial creando otro administrador o modificando la cuenta desde la base de datos antes de publicar el sistema en Internet.

## Uso
1. Inicia sesión y crea una rendición.
2. Sube uno o varios comprobantes PDF.
3. Completa los datos del técnico y de la comisión.
4. Pulsa **Procesar rendición y generar Excel**.
5. Revisa los campos extraídos. La clasificación es sugerida y requiere verificación humana.
6. Si corriges un dato, pulsa **Guardar corrección**; el Excel se vuelve a generar.
7. Descarga el Excel desde el botón superior.

## Cuenta inicial y seguridad
- Usuario: `admin`
- Contraseña inicial: `Admin123!`
- La clave de sesión debe cambiarse mediante la variable de entorno `SECRET_KEY` antes de cualquier despliegue público.
- No publiques esta versión directamente en Internet sin configurar HTTPS, una clave secreta fuerte, copias de seguridad, protección CSRF, política de contraseñas y un servidor WSGI de producción.
- Los archivos PDF y Excel se guardan localmente en `uploads/` y `exports/`; la base de datos es `zonda.db`.

## Alcance y limitaciones reales
- El extractor procesa PDFs que contienen texto seleccionable. Los PDF escaneados como imagen necesitan OCR adicional; en esta versión se reportará un error o faltarán campos para revisión, no se inventarán datos.
- La extracción se basa en patrones de texto del motor original. Comprobantes con formatos distintos pueden requerir corrección manual.
- La clasificación se basa en palabras clave del detalle y es orientativa.
- La plantilla conserva la estructura y formato original. Las fórmulas se escriben en el Excel; el recálculo de fórmulas lo realiza Excel/LibreOffice al abrir el archivo.
- La cuenta bancaria y los datos personales se deben introducir y verificar en cada rendición.
- La plantilla original corresponde al formato de rendición proporcionado. Si la entidad modifica el formato, habrá que actualizar el motor.

## Estructura
- `app.py`: servidor Flask, usuarios, rendiciones, historial y descargas.
- `engine.py`: extracción y generación Excel.
- `engine/`: plantilla Excel original y configuración de referencia.
- `templates/`, `static/`: interfaz web.
- `uploads/`, `exports/`: se crean automáticamente.

## Ejecución manual
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
\n\n## Mejoras incorporadas\n- Eliminación de rendiciones y archivos asociados.\n- Número de rendición correlativo automático.\n- Fecha de transferencia con selector de fecha.\n- Fecha y lugar de firma actualizados automáticamente al día actual.\n- Ubicación como campo editable sin texto de marca de agua.\n- Texto explicativo y título de movilidad generados automáticamente en el Excel; ya no se solicitan en el formulario.\n\n- Campo editable «Detalle del gasto», ubicado en la columna de detalle de la tabla Excel.\n- Retirado el texto de ejemplo/placeholder del campo Ubicación.\n
- Importe recibido validado como número entero, sin decimales.
- Logotipo MISTRAL incorporado en la esquina superior izquierda de la plantilla Excel.
- Texto explicativo automático con estructura de viaje, motivo y referencia a RENIEC.


## Dashboard de contabilidad
El menú **Contabilidad** está disponible para administradores. Incluye gastos totales, importes recibidos, saldo referencial, distribución por categorías, resumen mensual y rendiciones recientes. La clasificación por categoría depende del texto de detalle extraído y debe revisarse cuando aparezcan importes en «Otros».

## Publicación en Render desde GitHub
1. Sube el contenido de esta carpeta al repositorio privado de GitHub.
2. En Render, elige **New + > Blueprint** y conecta el repositorio.
3. Render leerá `render.yaml`. La aplicación usa un disco persistente montado en `/var/data` para la base SQLite, PDF y Excel.
4. Al crear el servicio, configura `INITIAL_ADMIN_PASSWORD` con una contraseña única de al menos 12 caracteres. `SECRET_KEY` se genera automáticamente.
5. Render asignará una URL HTTPS al servicio. El disco persistente requiere un plan de pago; revisa el precio vigente antes de confirmar.
6. Haz copias de seguridad periódicas del contenido de `/var/data`. No elimines ni recrees el disco sin respaldo.

El repositorio debe mantenerse privado. No subas la base de datos, comprobantes, archivos Excel generados, claves ni contraseñas. Antes de usar datos reales en producción, se recomienda añadir protección CSRF y copias de seguridad automatizadas.
