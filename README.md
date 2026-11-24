# Generación de Estudios Previos

Aplicativo exclusivo de Estudios Previos. El formulario, Excel, motor Word, tablas y almacenamiento existentes se conservan.

- App: https://extractestudiosprevios-hvpueverxg24qsx7z4e7ad.streamlit.app/
- Entrada: `streamlit_app.py`.
- Plantillas: `Plantilla.docx` y `PlantillaCargaMasivaEstudiosPrevios.xlsx`.
- Supabase: exclusivamente `zwpmdgpywgypdmqygsvv`.
- Tablas existentes: `EstudiosPrevios`, `DocumentosGenerados`, `BitacoraAuditoria`.
- Bucket existente: `documentos-generados` (conservar su configuración).
- Minutas se mantiene en https://github.com/Camilalozano/GeneracionMinutas, con despliegue propio.

## Base histórica

Se recuperó el aplicativo de `7ebb910e1232d94e72183f560f3fd479e6d05f45`, anterior a la integración de Minutas en `7bed8b9`. Incluye el ajuste para claves secret de Supabase. No se reescribe el historial Git.

## Ejecutar

La respuesta a licencias es siempre **No** en formulario, vista previa, Word y nuevos registros guardados. El Excel de carga masiva no incluye `requiere_licencias_si` ni `requiere_licencias_no`; si se carga un archivo antiguo, esas columnas se ignoran. Las obligaciones específicas se conservan.

La plantilla Excel es una tabla de captura sin conexión de actualización externa, para evitar que una actualización vuelva a incorporar las columnas retiradas. Se conservan las demás columnas, filas y formatos. Los campos internos de licencias se mantienen para completar las plantillas Word existentes y guardar la respuesta fija, sin modificar el histórico.

La fecha de elaboración (`{{hoy}}`) se completa en el Word descargable con la fecha del formulario. En carga masiva, si la columna `hoy` no existe o está vacía, se usa la fecha actual de Bogotá. Se admite una fecha explícita en formato DD/MM/AAAA, AAAA-MM-DD o una celda de fecha Excel. Una fecha inválida detiene la carga con un aviso.

```sh
pip install -r requirements.txt
streamlit run streamlit_app.py
python -m unittest discover -s tests -v
```

Conservar los secretos actuales de Streamlit: `SUPABASE_URL`, `SUPABASE_PUBLISHABLE_KEY`, `SUPABASE_SERVICE_ROLE_KEY`, `SUPABASE_BUCKET_DOCUMENTOS`, `ADMIN_PASSWORD`. Nunca subir `.streamlit/secrets.toml`.

## Datos y reversión

Esta separación no elimina, mueve ni modifica registros históricos ni archivos almacenados. Las nuevas minutas usan una tabla y bucket propios en el mismo proyecto; no es una base física distinta ni aislamiento de credenciales por tabla. El histórico combinado queda conservado donde está.
Para revertir el código, revertir el commit de separación. No borrar tablas ni buckets.
