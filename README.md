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

```sh
pip install -r requirements.txt
streamlit run streamlit_app.py
python -m unittest discover -s tests -v
```

Conservar los secretos actuales de Streamlit: `SUPABASE_URL`, `SUPABASE_PUBLISHABLE_KEY`, `SUPABASE_SERVICE_ROLE_KEY`, `SUPABASE_BUCKET_DOCUMENTOS`, `ADMIN_PASSWORD`. Nunca subir `.streamlit/secrets.toml`.

## Datos y reversión

Esta separación no elimina, mueve ni modifica registros históricos ni archivos almacenados. Las nuevas minutas usan una tabla y bucket propios en el mismo proyecto; no es una base física distinta ni aislamiento de credenciales por tabla. El histórico combinado queda conservado donde está.
Para revertir el código, revertir el commit de separación. No borrar tablas ni buckets.
