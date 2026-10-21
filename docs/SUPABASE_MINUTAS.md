# Complemento de minutas — esquema propuesto

Proyecto inspeccionado: `zwpmdgpywgypdmqygsvv`, el 20 de octubre de 2026.
Este SQL es una propuesta revisable: **no se ha aplicado a producción**.
La generación y descarga de los dos Word funciona sin aplicarlo; el guardado
adicional de minutas requiere esta tabla y la clave de servidor ya usada para Storage.

## Estado encontrado

- `EstudiosPrevios`: datos del estudio, `id` identity. Se conserva sin cambios.
- `DocumentosGenerados`: metadatos de archivos; `id_caso` es una identity de documento,
  no una referencia al estudio previo. Se conserva el guardado existente del estudio.
- `BitacoraAuditoria`: se conserva.
- Bucket `documentos-generados`: ya existe y es privado. Se reutiliza con prefijo `minutas/`.

## Tabla adicional

Cada generación recibe un UUID. Guarda la instantánea exacta del estudio utilizada para
producir ambos Word, su hash y su ruta, además de los campos de la minuta y la versión
de la plantilla. No se relaciona por `codigo_objeto`, que no identifica una versión única.
Se conserva una instantánea porque el flujo actual inserta estudios con `return=minimal`
y no devuelve su id; cambiarlo alteraría el flujo que se pidió preservar.

Ejecutar una sola vez desde SQL Editor, dentro de una operación registrada de despliegue:

```sql
begin;
create table public."MinutasGeneradas" (
  id uuid primary key,
  created_at timestamptz not null default now(),
  numero_contrato text not null check (numero_contrato ~ '^[0-9]+$'),
  anio_contrato integer not null check (anio_contrato between 2000 and 2099),
  actor text,
  version_plantilla text not null,
  plantilla_sha256 text not null,
  datos_minuta jsonb not null,
  estudio_previo_snapshot jsonb not null,
  estudio_previo_sha256 text not null,
  estudio_previo_storage_path text,
  storage_bucket text not null,
  storage_path text not null unique,
  minuta_sha256 text not null
);
alter table public."MinutasGeneradas" enable row level security;
revoke all on public."MinutasGeneradas" from anon, authenticated;
grant select, insert on public."MinutasGeneradas" to service_role;
create index minutas_contrato_idx on public."MinutasGeneradas" (anio_contrato, numero_contrato);
notify pgrst, 'reload schema';
commit;
```

No se crean políticas públicas: la aplicación utiliza su clave de servidor y el panel
administrativo existente. No se agregan datos de identidad a `EstudiosPrevios`, cuyas
políticas actuales permiten lectura e inserción anónima. Este complemento no cambia
esas políticas ni incorpora autenticación nueva.

## Configuración de Streamlit

Conservar `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`, `SUPABASE_BUCKET_DOCUMENTOS`
y `ADMIN_PASSWORD` únicamente en Secrets. No subir valores al repositorio.
La clave puede ser una clave secret nueva o service_role legacy; la ruta nueva admite ambas.

Opcionalmente agregar una sección `[MINUTA_INSTITUCION]` con los valores verificados de
los campos de `INSTITUCION` en `minutas.py`. El formulario permite revisarlos para cada
contrato. No se precargan nombres ni identificaciones reales desde los PDF.
Sin esa sección, los campos se diligencian manualmente; solo se sugieren código,
versión y fecha de aprobación observados en el formato 2026.

## Comprobación tras aplicar

1. Generar un par con datos de prueba y comprobar ambas descargas.
2. Verificar en el panel administrativo una fila en `MinutasGeneradas`, su snapshot,
   dos hashes y rutas a los Word. Una ruta de estudio vacía significa que falló su
   guardado anterior, aunque ambos archivos se descargan.
3. Verificar que `anon` y `authenticated` no puedan leer ni escribir la tabla nueva.
4. El guardado consulta primero la tabla: si falta, no sube una minuta huérfana.
   Si la red falla después de subir el archivo, puede quedar un objeto sin registro;
   revisar el prefijo `minutas/<UUID>/` antes de reintentar. No se anuncia guardado exitoso.

No se ha ejecutado una generación con datos reales en producción.

Referencias: [claves API](https://supabase.com/docs/guides/api/api-keys),
[RLS](https://supabase.com/docs/guides/database/postgres/row-level-security).
