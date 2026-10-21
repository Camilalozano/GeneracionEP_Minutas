# Generación de estudio previo y minuta

La sección 8 complementa el formulario existente con número/año contractual,
identificación de la persona natural, CDP y datos institucionales. Los campos
contratista, tipo de contrato, código de objeto y objeto se reutilizan sin duplicar captura.

Las secciones 1 a 7, sus validaciones, `Plantilla.docx`, la plantilla Excel y el motor
de estudios previos permanecen iguales. El tipo de contrato se normaliza únicamente
para el texto de la minuta: servicios profesionales o apoyo a la gestión.

1. Completar el formulario, incluida la sección 8, y pulsar **Usar este registro**.
2. Revisar las vistas previas descargables de ambos Word.
3. Pulsar **Generar Documentos**. El ZIP contiene una carpeta por registro con
   `Estudio_previo.docx` y `Minuta.docx`. No se genera un par incompleto si falta un dato.

La vista previa de estudios previos sigue disponible aunque falten datos de la minuta.
El Excel anterior sigue funcionando en su modo habitual; para generar pares en lote,
activar **Generar también una minuta por cada fila del Excel** y agregar las columnas
indicadas en pantalla. Identificaciones, consecutivos y años deben llegar como texto
sin decimales. No se cambia la plantilla de carga masiva existente.

## Plantilla

`templates/01_Plantilla_minuta_persona_natural.docx` conserva el contenido y formato de la plantilla
revisada con Camila. Se retiraron únicamente los metadatos personales de autoría. Conserva las 31 cláusulas y las remisiones al estudio previo.
No se incorpora la variante de persona jurídica. Los campos se reemplazan también
en tablas y encabezados, preservando formato y campos de número de página. Las
variables azules se convierten a texto negro en el documento diligenciado.

El año del contrato se captura explícitamente. Las fechas de los actos administrativos
y del formato no se recalculan. Deben verificarse las referencias institucionales antes
de usar el documento. Se conservan las observaciones editoriales de la plantilla base.

## Pruebas

```sh
python -m unittest discover -s tests -v
```

Incluyen variables repartidas entre runs, encabezados y tablas, datos obligatorios,
fechas inválidas, generación de dos pares, fallos de almacenamiento, preservación de
descargas y regresión del motor, formulario y plantilla de estudios previos frente a `main`.
La persistencia nueva y el SQL pendiente se documentan en [SUPABASE_MINUTAS.md](SUPABASE_MINUTAS.md).
