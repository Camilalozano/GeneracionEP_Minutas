"""Complemento de minutas; no modifica el motor ni la plantilla de estudios previos."""
import hashlib
import io
import re
import unicodedata
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document

TEMPLATE = Path(__file__).parent / 'templates/01_Plantilla_minuta_persona_natural.docx'
VERSION = 'persona-natural-2026-v1'
CONTRATO = {
    'numero_contrato': 'Número de contrato (solo consecutivo)',
    'anio_contrato': 'Año del contrato',
    'contratista_tipo_documento': 'Tipo de documento del contratista',
    'contratista_documento': 'Número de documento del contratista',
    'cdp_numero': 'Número del CDP',
    'cdp_fecha': 'Fecha del CDP (día/mes/año)',
    'cdp_expedidor': 'Cargo de quien expide el CDP',
}
INSTITUCION = {
    'formato_codigo': 'Código del formato',
    'formato_version': 'Versión del formato',
    'formato_fecha_aprobacion': 'Fecha de aprobación del formato (día/mes/año)',
    'entidad_nombre': 'Nombre de la entidad contratante',
    'entidad_nit': 'NIT de la entidad contratante',
    'ordenador_nombre': 'Nombre del ordenador del gasto',
    'ordenador_tipo_documento': 'Tipo de documento del ordenador',
    'ordenador_documento': 'Número de documento del ordenador',
    'ordenador_cargo': 'Cargo del ordenador',
    'acto_delegacion': 'Acto de delegación (referencia completa)',
    'acto_nombramiento': 'Acto de nombramiento (referencia completa)',
    'acto_posesion': 'Acta de posesión (referencia completa)',
}
COMPARTIDOS = ('contratista', 'tipo_contrato', 'codigo_objeto', 'objeto_contrato')
DEFAULTS = {'formato_codigo': 'F7_P11_C', 'formato_version': '2',
            'formato_fecha_aprobacion': '09/12/2025',
            'contratista_tipo_documento': 'cédula de ciudadanía'}
MESES = ('enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre').split()


def texto(value):
    return '' if value is None or pd.isna(value) else str(value).strip()


def modalidad(value):
    normalized = ''.join(c for c in unicodedata.normalize('NFKD', texto(value).lower())
                         if not unicodedata.combining(c))
    normalized = ' '.join(normalized.split())
    accepted = {
        'servicios profesionales': 'PRESTACIÓN DE SERVICIOS PROFESIONALES',
        'prestacion de servicios profesionales': 'PRESTACIÓN DE SERVICIOS PROFESIONALES',
        'apoyo a la gestion': 'PRESTACIÓN DE SERVICIOS DE APOYO A LA GESTIÓN',
        'servicios de apoyo a la gestion': 'PRESTACIÓN DE SERVICIOS DE APOYO A LA GESTIÓN',
        'prestacion de servicios de apoyo a la gestion': 'PRESTACIÓN DE SERVICIOS DE APOYO A LA GESTIÓN',
    }
    if normalized.startswith('contrato de '):
        normalized = normalized[len('contrato de '):]
    return accepted.get(normalized)


def validar_minuta(data):
    errors = []
    for key in (*COMPARTIDOS, *CONTRATO, *INSTITUCION):
        value = texto(data.get(key))
        if not value:
            errors.append(f"{CONTRATO.get(key, INSTITUCION.get(key, key))}: falta diligenciar.")
        elif '{{' in value or '}}' in value:
            errors.append(f'{key}: reemplaza las llaves por el dato real.')
    if texto(data.get('numero_contrato')) and not re.fullmatch(r'\d+', texto(data['numero_contrato'])):
        errors.append('Número de contrato: escribe solo el consecutivo, por ejemplo 001.')
    if texto(data.get('anio_contrato')) and not re.fullmatch(r'20\d{2}', texto(data['anio_contrato'])):
        errors.append('Año del contrato: usa cuatro dígitos, por ejemplo 2026.')
    if texto(data.get('tipo_contrato')) and not modalidad(data['tipo_contrato']):
        errors.append('La minuta admite servicios profesionales o servicios de apoyo a la gestión.')
    for key in ('cdp_fecha', 'formato_fecha_aprobacion'):
        if texto(data.get(key)):
            try:
                datetime.strptime(texto(data[key]), '%d/%m/%Y')
            except ValueError:
                errors.append(f'{key}: ingresa una fecha válida en formato día/mes/año.')
    return errors


def formulario_minuta(st, configuracion=None):
    st.markdown('#### 8) Minuta de persona natural')
    st.caption('Se generarán el estudio previo y la minuta con el mismo contratista y objeto. '
               'Completa estos datos y pulsa «Usar este registro» para aplicarlos.')
    values = {}
    defaults = {**DEFAULTS, **(configuracion or {})}
    for group, fields in [('Contrato y disponibilidad presupuestal', CONTRATO),
                          ('Datos institucionales vigentes para este contrato', INSTITUCION)]:
        with st.expander(group, expanded=group.startswith('Contrato')):
            columns = st.columns(2)
            for i, (key, label) in enumerate(fields.items()):
                with columns[i % 2]:
                    values[key] = st.text_input(label + ' *', value=texto(defaults.get(key)), key='minuta_' + key)
    st.caption('Las fechas de delegación, nombramiento, posesión y aprobación del formato '
               'no cambian automáticamente con el año del contrato.')
    return values


def partes(doc):
    yield doc.element
    # Existing parts only: accessing undefined first/even headers creates new parts.
    for part in doc.part.related_parts.values():
        element = getattr(part, 'element', None)
        if element is not None and element.tag.rsplit('}', 1)[-1] in ('hdr', 'ftr'):
            yield element


def generar_minuta(data, plantilla=TEMPLATE):
    errors = validar_minuta(data)
    if errors:
        raise ValueError('\n'.join(errors))
    values = {k: texto(data.get(k)) for k in (*COMPARTIDOS, *CONTRATO, *INSTITUCION)}
    values['tipo_contrato'] = modalidad(values['tipo_contrato'])
    fecha = datetime.strptime(values['cdp_fecha'], '%d/%m/%Y')
    values['cdp_fecha'] = f'{fecha.day} de {MESES[fecha.month - 1]} de {fecha.year}'
    doc = Document(plantilla)
    # Replace right-to-left across text nodes: preserves formatting, tables, fields and headers.
    for part in partes(doc):
        for paragraph in part.xpath('.//w:p'):
            nodes = paragraph.xpath('.//w:t')
            full = ''.join(n.text or '' for n in nodes)
            spans, pos = [], 0
            for n in nodes:
                spans.append((pos, pos + len(n.text or ''), n))
                pos = spans[-1][1]
            for match in reversed(list(re.finditer(r'\{\{([a-z_]+)\}\}', full))):
                key = match.group(1)
                if key not in values:
                    raise ValueError(f'Variable desconocida en la minuta: {key}')
                affected = [(a,b,n) for a,b,n in spans if a < match.end() and b > match.start()]
                for i, (a,b,n) in enumerate(affected):
                    before = (n.text or '')[:max(0, match.start()-a)]
                    after = (n.text or '')[max(0, match.end()-a):] if match.end() < b else ''
                    n.text = before + (values[key] if i == 0 else '') + after
                    n.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            if '{{' in ''.join(n.text or '' for n in nodes):
                raise ValueError('La minuta contiene variables sin resolver.')
    # Blue identifies fields in the reusable template, not in the completed contract.
    for part in partes(doc):
        for color in part.xpath('.//w:rPr/w:color'):
            if color.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '').upper() == '155D9C':
                color.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '000000')
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def generar_pares(df, word_file, progress, status, actor, generar_ep, guardar_ep,
                  guardar_minuta=None):
    """Build both files before including a row; persistence failures never remove downloads."""
    output, count, saved, errors = io.BytesIO(), 0, 0, []
    for position, (_, row) in enumerate(df.iterrows(), 1):
        for error in validar_minuta(row):
            errors.append(f'Fila {position}: {error}')
    if errors:
        return output, count, errors, saved
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        for position, (_, row) in enumerate(df.iterrows(), 1):
            try:
                minuta = generar_minuta(row)
                ep = io.BytesIO()
                generar_ep(word_file, row).save(ep)
                ep_bytes = ep.getvalue()
                pair_id = str(uuid.uuid4())
                folder = f'{position:03d}_ATENEA-{texto(row["numero_contrato"])}-{texto(row["anio_contrato"])}'
                archive.writestr(folder + '/Estudio_previo.docx', ep_bytes)
                archive.writestr(folder + '/Minuta.docx', minuta)
                count += 2
                ep_metadata = None
                if guardar_ep:
                    try:
                        ep_metadata = guardar_ep(row, position-1, 'Estudio_previo.docx', ep_bytes, actor)
                        saved += 1
                    except Exception:
                        errors.append(f'Fila {position}: el estudio previo se descargará, pero no se pudo guardar en Supabase.')
                if guardar_minuta:
                    try:
                        guardar_minuta(row, pair_id, minuta, ep_bytes, ep_metadata, actor)
                        saved += 1
                    except Exception:
                        errors.append(f'Fila {position}: la minuta se descargará, pero no se pudo registrar en Supabase. Revisa la configuración de minutas y el esquema adicional.')
            except Exception as exc:
                errors.append(f'Fila {position}: no se generó el par de documentos: {exc}')
            progress.progress(position / len(df))
            status.text(f'Procesando par {position} de {len(df)}...')
    status.text('Proceso completado.')
    output.seek(0)
    return output, count, errors, saved


def guardar_minuta_supabase(row, pair_id, minuta, ep, ep_metadata, actor, *,
                            url, key, bucket, campos_ep, request):
    """New private registry. Does not write to or change EstudiosPrevios."""
    from urllib.parse import quote
    if not key:
        raise ValueError('Falta la clave de servidor para guardar minutas.')
    headers = {'apikey': key}
    if key.startswith('eyJ'):
        headers['Authorization'] = 'Bearer ' + key
    # Check schema before uploading to avoid an orphan when the proposed SQL is not applied.
    check = request('GET', url + '/rest/v1/MinutasGeneradas', headers=headers,
                    params={'select': 'id', 'limit': '0'}, timeout=20)
    check.raise_for_status()
    path = f'minutas/{pair_id}/Minuta.docx'
    endpoint = f'{url}/storage/v1/object/{quote(bucket, safe="")}/{path}'
    upload = request('POST', endpoint, headers={**headers, 'Content-Type':
                     'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                     'x-upsert': 'false'}, data=minuta, timeout=45)
    upload.raise_for_status()
    record = {
        'id': pair_id, 'numero_contrato': texto(row['numero_contrato']),
        'anio_contrato': int(texto(row['anio_contrato'])), 'actor': actor,
        'version_plantilla': VERSION, 'plantilla_sha256': hashlib.sha256(TEMPLATE.read_bytes()).hexdigest(),
        'datos_minuta': {k: texto(row.get(k)) for k in (*COMPARTIDOS, *CONTRATO, *INSTITUCION)},
        'estudio_previo_snapshot': {k: texto(row.get(k)) for k in campos_ep},
        'estudio_previo_sha256': hashlib.sha256(ep).hexdigest(),
        'estudio_previo_storage_path': (ep_metadata or {}).get('storage_path'),
        'storage_bucket': bucket, 'storage_path': path,
        'minuta_sha256': hashlib.sha256(minuta).hexdigest(),
    }
    response = request('POST', url + '/rest/v1/MinutasGeneradas', headers={**headers,
                       'Content-Type': 'application/json', 'Prefer': 'return=minimal'}, json=record, timeout=20)
    response.raise_for_status()
    return record
