import ast
import io
from pathlib import Path
import subprocess
import unittest
import zipfile
from unittest.mock import patch
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]


class EstudiosTests(unittest.TestCase):
    def test_motor_y_plantilla_historicos(self):
        old = subprocess.check_output(['git', 'show', '7ebb910:streamlit_app.py'], cwd=ROOT).decode('utf-8')
        current = (ROOT/'streamlit_app.py').read_text(encoding='utf-8')
        for name in ['generar_documentos', 'supabase_request',
                     'guardar_estudios_previos_supabase']:
            nodes = [next(n for n in ast.parse(s).body if isinstance(n, ast.FunctionDef) and n.name == name)
                     for s in [old, current]]
            self.assertEqual(ast.dump(nodes[0]), ast.dump(nodes[1]))
        self.assertNotIn('from minutas', current)
        self.assertEqual((ROOT/'Plantilla.docx').read_bytes(), subprocess.check_output(
            ['git', 'show', '7ebb910:Plantilla.docx'], cwd=ROOT))

    @patch('requests.request', side_effect=AssertionError('Sin red en pruebas'))
    def test_formulario_genera_solo_estudio(self, _):
        app = AppTest.from_file(str(ROOT/'streamlit_app.py'), default_timeout=40)
        for key in ['SUPABASE_URL', 'SUPABASE_PUBLISHABLE_KEY', 'SUPABASE_SERVICE_ROLE_KEY']:
            app.secrets[key] = ''
        app.run()
        self.assertFalse(app.exception)
        self.assertFalse(any('Requiere licencias' in radio.label for radio in app.radio))
        for field in app.text_input:
            if field.label.endswith('*') or field.label == 'Contratista':
                field.set_value('Dato de prueba')
        next(f for f in app.text_input if f.label == 'Número de proyecto *').set_value('123')
        for field in app.text_area:
            if not field.disabled:
                field.set_value('Texto de prueba para completar el estudio previo aprobado.')
        next(b for b in app.button if b.label == '✅ Usar este registro').click().run()
        next(b for b in app.button if b.label == '🚀 Generar Estudios Previos').click().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.session_state['resultado_generados'], 1)
        with zipfile.ZipFile(io.BytesIO(app.session_state['resultado_zip'])) as z:
            self.assertEqual(len(z.namelist()), 1)
            self.assertNotIn('Minuta', z.namelist()[0])

    def test_licencias_fijas_en_word_y_registros(self):
        import pandas as pd
        from docx import Document
        source = ast.parse((ROOT/'streamlit_app.py').read_text(encoding='utf-8'))
        names = {'LICENCIAS_FIJAS', 'PLANTILLA_VARIABLES', 'SUPABASE_COLUMNAS_ESTUDIOS',
                 'SUPABASE_COLUMNAS_ENTERAS', 'SUPABASE_COLUMNAS_FECHA'}
        functions = {'replace_text_in_paragraph', 'diligenciar_documento',
                     'normalizar_valor_supabase', 'normalizar_registro_estudios'}
        nodes = [n for n in source.body if
                 (isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id in names for t in n.targets))
                 or (isinstance(n, ast.FunctionDef) and n.name in functions)]
        namespace = {'pd': pd, 'Document': Document}
        exec(compile(ast.Module(body=nodes, type_ignores=[]), '<test>', 'exec'), namespace)
        for data in [{}, {'requiere_licencias_si': 'X', 'requiere_licencias_no': ''}]:
            with self.subTest(data=data):
                template = Document()
                template.add_paragraph('Sí: {{requiere_licencias_si}} / No: {{requiere_licencias_no}}')
                template.add_table(rows=1, cols=1).cell(0, 0).text = '{{requiere_licencias_no}}'
                buffer = io.BytesIO()
                template.save(buffer)
                document = namespace['diligenciar_documento'](buffer, pd.Series(data, dtype=object))
                self.assertEqual(document.paragraphs[0].text, 'Sí:  / No: X')
                self.assertEqual(document.tables[0].cell(0, 0).text, 'X')
                record = namespace['normalizar_registro_estudios'](data)
                self.assertIsNone(record['requiere_licencias_si'])
                self.assertEqual(record['requiere_licencias_no'], 'X')

    def test_excel_sin_columnas_licencias(self):
        import openpyxl
        workbook = openpyxl.load_workbook(ROOT/'PlantillaCargaMasivaEstudiosPrevios.xlsx')
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        self.assertNotIn('requiere_licencias_si', headers)
        self.assertNotIn('requiere_licencias_no', headers)
        table = next(iter(sheet.tables.values()))
        self.assertEqual([col.name for col in table.tableColumns], headers)
        self.assertEqual(sheet.max_column, 26)
