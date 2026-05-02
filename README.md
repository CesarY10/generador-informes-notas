<h1 align="center"> Generador Automático de Informes de Notas</h1>

<p align="center">
Automatización con Python para generar informes académicos en Word a partir de datos en Excel.
<br>
Validación de datos • Cálculo de notas • Documentos personalizados
</p>

<p align="center">
<img src="https://img.shields.io/badge/Python-3.10-blue?style=for-the-badge">
<img src="https://img.shields.io/badge/pandas-Data%20Processing-yellow?style=for-the-badge">
<img src="https://img.shields.io/badge/docxtpl-Word%20Automation-green?style=for-the-badge">
<img src="https://img.shields.io/badge/openpyxl-Excel%20Engine-orange?style=for-the-badge">
</p>

---

##  Tabla de Contenidos

- [Descripción del Proyecto](#-descripción-del-proyecto)
- [Características](#-características)
- [Tecnologías Utilizadas](#-tecnologías-utilizadas)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Flujo del Programa](#-flujo-del-programa)
- [Cómo Ejecutarlo](#-cómo-ejecutarlo)
- [Ejemplo de Salida](#-ejemplo-de-salida)
- [Posibles Mejoras](#-posibles-mejoras)
- [Aprendizajes Clave](#-aprendizajes-clave)
- [Autor](#-autor)

---

##  Descripción del Proyecto

Este proyecto automatiza la generación de informes académicos en formato Word a partir de datos almacenados en un archivo Excel.  
El sistema:

- Valida la información de entrada  
- Calcula promedios y calificaciones finales  
- Asigna colores según el rendimiento  
- Rellena una plantilla Word personalizada por alumno  

Es ideal para instituciones educativas, automatización administrativa o como demostración de habilidades en Python, manipulación de datos y generación dinámica de documentos.

---

##  Características

- Lectura de datos desde Excel con pandas  
- Validación automática de:
  - Asignaturas faltantes  
  - Asignaturas duplicadas  
  - Notas fuera de rango  
- Cálculo de nota final y calificación textual  
- Asignación de colores según rendimiento  
- Generación de documentos Word personalizados con docxtpl  
- Limpieza y creación automática de carpetas de salida  
- Flujo completamente automatizado por alumno  

---

##  Tecnologías Utilizadas

- **Python 3.10**
- **pandas** — procesamiento de datos  
- **openpyxl** — motor para Excel  
- **docxtpl** — plantillas Word dinámicas  
- **shutil / os** — manejo de archivos y carpetas  

---

##  Estructura del Proyecto

```text
generador-informes-notas/
│
├── inputs/
│   ├── Notas_Alumnos.xlsx
│   └── Plantilla_Notas.docx
│
├── outputs/
│   └── .gitkeep
│
├── main.py
├── requirements.txt
└── README.md

