"""Prueba del formulario real con datos ficticios y sin conexiones de producción."""
import io
import unittest
import zipfile
from unittest.mock import patch

try:
    from streamlit.testing.v1 import AppTest
except ImportError:
    AppTest = None

from test_minutas import fixture, ROOT


@unittest.skipIf(AppTest is None, 'Instalar Streamlit para ejecutar la prueba de interfaz')
class FormularioTest(unittest.TestCase):
    @patch('requests.request', side_effect=AssertionError('No se permiten conexiones en esta prueba'))
    @patch('requests.post', side_effect=AssertionError('No se permiten conexiones en esta prueba'))
    def test_formulario_genera_dos_documentos(self, *_):
        app = AppTest.from_file(str(ROOT/'streamlit_app.py'), default_timeout=40)
        app.secrets['SUPABASE_URL'] = ''
        app.secrets['SUPABASE_PUBLISHABLE_KEY'] = ''
        app.secrets['SUPABASE_SERVICE_ROLE_KEY'] = ''
        app.run()
        self.assertFalse(app.exception, [e.message for e in app.exception])
        values = fixture()
        for field in app.text_input:
            if field.key and field.key.startswith('minuta_'):
                field.set_value(values[field.key[len('minuta_'):]])
            elif field.label.endswith('*') or field.label == 'Contratista':
                field.set_value('Dato de prueba')
        next(f for f in app.text_input if f.label == 'Tipo de contrato *').set_value(values['tipo_contrato'])
        next(f for f in app.text_input if f.label == 'Número de proyecto *').set_value('123')
        for field in app.text_area:
            if not field.disabled:
                field.set_value('Texto de prueba para completar el estudio previo aprobado.')
        next(b for b in app.button if b.label == '✅ Usar este registro').click().run()
        self.assertFalse(app.exception, [e.message for e in app.exception])
        self.assertTrue(any(d.label == 'Descargar vista previa de la minuta (.docx)' for d in app.get('download_button')))
        next(b for b in app.button if b.label == '🚀 Generar Documentos').click().run()
        self.assertFalse(app.exception, [e.message for e in app.exception])
        self.assertEqual(app.session_state['resultado_generados'], 2)
        with zipfile.ZipFile(io.BytesIO(app.session_state['resultado_zip'])) as package:
            self.assertEqual(len(package.namelist()), 2)
            self.assertTrue(any(n.endswith('/Minuta.docx') for n in package.namelist()))
        # A fresh submitted capture must not leave the old ZIP visible.
        app.text_input(key='minuta_cdp_numero').set_value('')
        next(b for b in app.button if b.label == '✅ Usar este registro').click().run()
        self.assertIsNone(app.session_state['resultado_zip'])
        next(b for b in app.button if b.label == '🚀 Generar Documentos').click().run()
        self.assertIsNone(app.session_state['resultado_zip'])
        self.assertTrue(any('cdp' in w.value.lower() for w in app.warning))


if __name__ == '__main__':
    unittest.main()
