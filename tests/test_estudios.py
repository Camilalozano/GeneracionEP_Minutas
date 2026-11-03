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
        for name in ['diligenciar_documento', 'generar_documentos', 'supabase_request',
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
