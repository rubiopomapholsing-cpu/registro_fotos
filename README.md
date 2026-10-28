# Registro fotográfico de fiscalización

Aplicación web en HTML, CSS y JavaScript clásico para organizar fotografías por proyecto, asignar descripciones y generar un anexo fotográfico real en Word. Los proyectos, el banco de descripciones y las imágenes se conservan localmente en IndexedDB.

## Ejecución

1. Abre esta carpeta en Visual Studio Code.
2. Inicia Live Server desde `index.html` o sirve la carpeta con un servidor web estático.
3. Abre la dirección local mostrada por el servidor. No hay un proceso de compilación ni dependencias de ejecución que instalar.

La biblioteca `docx` v8.5.0 está incluida localmente en `js/vendor/docx.umd.js`; no se descarga nada desde un CDN al abrir la aplicación. Para regenerar la copia, instala `docx@8.5.0` en una carpeta temporal con npm y copia `node_modules/docx/build/index.umd.js` a `js/vendor/docx.umd.js`.

## Uso

1. Completa título, razón social o establecimiento, código, fecha, lugar y fiscalizador. Cada campo se guarda automáticamente al salir; **Guardar proyecto** también guarda los datos actuales.
2. Pulsa **Tomar foto** o **Agregar imágenes**. Revisa las vistas previas, añade opcionalmente una descripción inicial y confirma o descarta antes de guardar.
3. Selecciona una foto pendiente. Elige una descripción disponible para asignarla, o escribe una propia. **+ Agregar descripción** permite pegar una descripción por línea o cargar un archivo `.txt`.
4. En una descripción puedes usar `{RAZON_SOCIAL}` y `{PLACAS}`. Se sustituyen con los valores del proyecto en la vista previa y el documento Word. Las descripciones asignadas dejan de aparecer en el banco disponible.
5. Filtra **Todas**, **Pendientes** o **Asignadas** y utiliza **Subir / Bajar** para cambiar la secuencia del Word. **Devolver a pendiente** libera la descripción para volver a usarla.
6. **Generar Word** requiere todos los campos obligatorios, al menos una foto y ninguna foto pendiente. La vista previa muestra las fotos asignadas en el mismo orden y con la misma numeración del documento.

## Captura desde un teléfono

- Conecta el teléfono y la computadora a la misma Wi-Fi. Inicia el servidor de desarrollo enlazado a todas las interfaces, por ejemplo `py -m http.server 8000 --bind 0.0.0.0 --directory "C:\ruta\al\proyecto"`.
- Consulta la IPv4 Wi-Fi de la computadora con `ipconfig` y abre `http://IP-DE-LA-COMPUTADORA:8000` desde el navegador del teléfono. Permite el acceso de Python a redes privadas en el Firewall de Windows si lo solicita.
- Para publicar la aplicación o acceder desde otra red, usa un alojamiento HTTPS. La cámara depende del selector de archivos/captura del navegador; su interfaz y disponibilidad varían entre dispositivos. Si cancelas o rechazas la captura, aparece un aviso y puedes continuar.
- Los datos se guardan en el origen y navegador usados en cada dispositivo. No se sincronizan entre el teléfono y la computadora ni se respaldan automáticamente.

## Estructura

- `index.html`: interfaz, cuadros de diálogo, plantilla de filas y scripts clásicos.
- `css/estilos.css`: distribución de escritorio y estilos adaptables para móvil.
- `js/app.js`: captura, eventos, guardado automático, asignaciones, validación y vista previa.
- `js/db.js`: proyectos, banco y fotos Blob en IndexedDB; migración de la base anterior.
- `js/imagenes.js`: corrección de orientación, reducción a 1600 px y conversión a JPEG.
- `js/exportar.js`: construcción y descarga del anexo `.docx`.
- `js/vendor/docx.umd.js`: biblioteca local `docx` v8.5.0.

## Eventos, persistencia y Word

Los botones de captura abren inputs de archivo estándar con `accept="image/*"`; la captura trasera se solicita mediante `capture="environment"`. La aplicación prepara cada imagen antes de presentarla para confirmación y guarda el Blob JPEG (no una ruta ni una URL temporal). Cada proyecto tiene su propio banco y sus fotos se consultan mediante el índice `projectId`.

Las operaciones de escritura esperan `transaction.oncomplete`. El reordenamiento y la eliminación/renumeración de fotos se guardan en una única transacción. Al exportar se genera un DOCX con página A4, márgenes de 2,54 cm, texto de 8 pt, imágenes centradas y filas que no se parten. Se inserta un salto antes de la tercera, quinta y siguientes fotos impares para maquetar dos fotos por página.

## Limitaciones conocidas

- IndexedDB es local al navegador y al origen; eliminar los datos del sitio elimina proyectos y fotos.
- La cámara real y la interfaz para capturar dependen del dispositivo y navegador; deben probarse desde el teléfono. El modo `capture` puede abrir la cámara o el selector de archivos.
- Las imágenes se recodifican como JPEG; la fecha que se registra es la de incorporación a la aplicación y no certifica cuándo se tomó originalmente la foto.
- La familia Poppins debe estar instalada en el equipo que abre Word para representarse exactamente; no se incrusta.
- El salto de página solicitado se incluye en el DOCX, pero Word y LibreOffice pueden ajustar el paginado cuando cambian fuentes o una descripción ocupa mucho espacio.
