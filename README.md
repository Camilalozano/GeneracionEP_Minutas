# Generador Inteligente de Documentos

Aplicación web para la generación automatizada de documentos Word a partir de datos en Excel. Utiliza plantillas parametrizadas con placeholders `{{variable}}` para crear múltiples documentos personalizados de manera eficiente.

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://automadedoc.streamlit.app/)

---

## Descripción

Esta herramienta permite automatizar la creación de documentos mediante:

- Carga de un archivo Excel (.xlsx) con datos a procesar
- Carga de una plantilla Word (.docx) con placeholders
- Vista previa de datos antes de generar
- Generación masiva de documentos con barra de progreso
- Descarga de documentos generados en formato ZIP

**Ideal para:** certificados, contratos, cartas, actas, informes y cualquier documento que necesite personalizarse masivamente.

---

## Demo en vivo

**Prueba la aplicación aquí:** [https://automadedoc.streamlit.app/](https://automadedoc.streamlit.app/)

---

## Tecnologías utilizadas

| Tecnología | Descripción |
|------------|-------------|
| Python 3.x | Lenguaje de programación |
| Streamlit | Framework para interfaz web |
| python-docx | Manipulación de documentos Word |
| pandas | Procesamiento de datos Excel |
| openpyxl | Motor para archivos .xlsx |

---

## Instalación local

### 1. Clona el repositorio

```bash
git clone https://github.com/tu-usuario/automotedoc.git
cd automotedoc
```

### 2. Instala las dependencias

```bash
pip install -r requirements.txt
```

### 3. Ejecuta la aplicación

```bash
streamlit run streamlit_app.py
```

La aplicación se abrirá en `http://localhost:8501`

---

## Cómo usar

### Paso 1: Prepara tu Excel

Crea un archivo Excel donde cada columna representa una variable y cada fila un documento a generar.

| nombre | cargo | fecha |
|--------|-------|-------|
| Juan Pérez | Analista | 2024-01-15 |
| María García | Coordinadora | 2024-01-16 |

### Paso 2: Prepara tu plantilla Word

Crea un documento Word usando placeholders con doble llave:

```
Certificamos que {{nombre}} desempeña el cargo de {{cargo}} desde el {{fecha}}.
```

### Paso 3: Genera los documentos

1. Sube ambos archivos en la aplicación
2. Haz clic en "Generar Documentos"
3. Descarga el ZIP con todos los documentos personalizados

---

## Estructura del proyecto

```
automotedoc/
├── streamlit_app.py    # Aplicación principal
├── requirements.txt    # Dependencias
├── README.md          # Este archivo
└── .gitignore         # Archivos ignorados
```

---

## Enlaces útiles

- [Documentación python-docx](https://python-docx.readthedocs.io/)
- [Documentación Streamlit](https://docs.streamlit.io/)
- [Aplicación desplegada](https://automadedoc.streamlit.app/)

