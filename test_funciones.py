import unittest
from unittest.mock import patch, MagicMock
from datetime import datetime
import os
import sys
import types

# Stubs de dependencias externas antes de importar 'funciones'
if 'qrcode' not in sys.modules:
    qrcode_stub = types.ModuleType('qrcode')
    qrcode_stub.make = MagicMock()
    sys.modules['qrcode'] = qrcode_stub

if 'fitz' not in sys.modules:
    fitz_stub = types.ModuleType('fitz')
    fitz_stub.open = MagicMock()
    class _Rect:
        def __init__(self, *args, **kwargs):
            pass
    fitz_stub.Rect = _Rect
    sys.modules['fitz'] = fitz_stub

if 'docxtpl' not in sys.modules:
    docxtpl_stub = types.ModuleType('docxtpl')
    class _DocxTemplate:
        def __init__(self, *args, **kwargs):
            pass
        def render(self, *args, **kwargs):
            pass
        def save(self, *args, **kwargs):
            pass
    class _RichText:
        def add(self, *args, **kwargs):
            pass
    docxtpl_stub.DocxTemplate = _DocxTemplate
    docxtpl_stub.RichText = _RichText
    sys.modules['docxtpl'] = docxtpl_stub

if 'docx2pdf' not in sys.modules:
    docx2pdf_stub = types.ModuleType('docx2pdf')
    def _convert(*args, **kwargs):
        pass
    docx2pdf_stub.convert = _convert
    sys.modules['docx2pdf'] = docx2pdf_stub

# docx.shared.Pt
if 'docx' not in sys.modules:
    docx_stub = types.ModuleType('docx')
    sys.modules['docx'] = docx_stub
if 'docx.shared' not in sys.modules:
    shared_stub = types.ModuleType('docx.shared')
    class Pt: pass
    shared_stub.Pt = Pt
    sys.modules['docx.shared'] = shared_stub

import funciones


class TestGenerarDocumento(unittest.TestCase):

    @patch('funciones.os.path.exists', return_value=False)
    def test_retora_none_si_plantilla_no_existe(self, mock_exists):
        fila = {
            'Detalle de servicio': 'Certificado de nivel',
            'Idioma': 'Español',
            'Nombres': 'Juan Perez',
            'Código': 'ABC123',
            'Indica el nivel culminado': 'intermedio',
            'Fecha examen o curso': datetime(2024, 5, 10),
        }
        resultado = funciones.generar_documento(fila)
        self.assertIsNone(resultado)

    @patch('funciones.os.remove')
    @patch('funciones.convert')
    @patch('funciones.DocxTemplate')
    @patch('funciones.os.path.exists', return_value=True)
    def test_exito_certificado_nivel(self, mock_exists, MockDocxTemplate, mock_convert, mock_remove):
        fila = {
            'Detalle de servicio': 'Certificado de nivel',
            'Idioma': 'Español',
            'Nombres': 'Juan Perez',
            'Código': 'ABC123',
            'Indica el nivel culminado': 'intermedio',
            'Fecha examen o curso': datetime(2024, 5, 10),
        }
        instancia_doc = MagicMock()
        MockDocxTemplate.return_value = instancia_doc

        resultado = funciones.generar_documento(fila)

        nombre_plantilla = 'Certificado_de_nivel_Español.docx'
        ruta_plantilla_esperada = os.path.join(funciones.config.RUTA_PLANTILLAS, nombre_plantilla)
        MockDocxTemplate.assert_called_once_with(ruta_plantilla_esperada)

        # Verificamos que render se llamó con un contexto con claves esperadas
        instancia_doc.render.assert_called_once()
        args, _ = instancia_doc.render.call_args
        contexto = args[0]
        self.assertIn('nombres_completos', contexto)
        self.assertIn('codigo_doc', contexto)
        self.assertIn('nivel_descripcion', contexto)
        self.assertIn('nivel_codigo', contexto)
        self.assertIn('idioma_texto', contexto)
        self.assertIn('examen_texto_simple', contexto)
        self.assertIn('emision_texto_simple', contexto)

        nombre_archivo = f"{fila['Código']} - {fila['Detalle de servicio']} - {fila['Idioma']} - {fila['Nombres']}"
        ruta_docx_esperada = os.path.join(funciones.config.RUTA_SALIDAS, f"{nombre_archivo}.docx")
        ruta_pdf_esperada = os.path.join(funciones.config.RUTA_SALIDAS, f"{nombre_archivo}.pdf")

        instancia_doc.save.assert_called_once_with(ruta_docx_esperada)
        mock_convert.assert_called_once_with(ruta_docx_esperada, ruta_pdf_esperada)
        mock_remove.assert_called_once_with(ruta_docx_esperada)
        self.assertEqual(resultado, ruta_pdf_esperada)

    @patch('funciones.os.remove')
    @patch('funciones.convert')
    @patch('funciones.DocxTemplate')
    @patch('funciones.os.path.exists', return_value=True)
    def test_exito_examen_comprension_textos(self, mock_exists, MockDocxTemplate, mock_convert, mock_remove):
        fila = {
            'Detalle de servicio': 'Examen de comprensión de textos',
            'Idioma': 'Inglés',
            'Nombres': 'Ana Gómez',
            'Código': 'XYZ789',
            'Resultado examen o curso': 'Aprobado',
            'Fecha examen o curso': datetime(2024, 7, 2),
        }
        instancia_doc = MagicMock()
        MockDocxTemplate.return_value = instancia_doc

        resultado = funciones.generar_documento(fila)

        nombre_plantilla = 'Examen_de_comprensión_de_textos_Inglés.docx'
        ruta_plantilla_esperada = os.path.join(funciones.config.RUTA_PLANTILLAS, nombre_plantilla)
        MockDocxTemplate.assert_called_once_with(ruta_plantilla_esperada)

        instancia_doc.render.assert_called_once()
        args, _ = instancia_doc.render.call_args
        contexto = args[0]
        # En este caso, nombres_completos debe ser texto simple y claves de fecha en texto_simple
        self.assertIsInstance(contexto['nombres_completos'], str)
        self.assertIn('examen_texto_simple', contexto)
        self.assertIn('emision_texto_simple', contexto)
        self.assertEqual(contexto['idioma_texto'], fila['Idioma'])

        nombre_archivo = f"{fila['Código']} - {fila['Detalle de servicio']} - {fila['Idioma']} - {fila['Nombres']}"
        ruta_docx_esperada = os.path.join(funciones.config.RUTA_SALIDAS, f"{nombre_archivo}.docx")
        ruta_pdf_esperada = os.path.join(funciones.config.RUTA_SALIDAS, f"{nombre_archivo}.pdf")

        instancia_doc.save.assert_called_once_with(ruta_docx_esperada)
        mock_convert.assert_called_once_with(ruta_docx_esperada, ruta_pdf_esperada)
        mock_remove.assert_called_once_with(ruta_docx_esperada)
        self.assertEqual(resultado, ruta_pdf_esperada)


if __name__ == '__main__':
    unittest.main()
