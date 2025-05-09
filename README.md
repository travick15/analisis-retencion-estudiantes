# Software de Análisis de Retención de Estudiantes

Este software analiza la retención de estudiantes a través de diferentes periodos académicos, identificando la continuidad de estudiantes nuevos en periodos subsecuentes.

## Requisitos

- Python 3.7 o superior
- pandas
- openpyxl

## Instalación

1. Clone o descargue este repositorio
2. Instale las dependencias:
```bash
pip install -r requirements.txt
```

## Estructura del Archivo de Datos

El software espera un archivo Excel con las siguientes columnas:
- "Periodo" (columna A): Identificador del periodo académico
- "Documento" (columna C): Número de documento del estudiante
- "Condición" (columna I): Condición del estudiante (debe incluir "nuevo" para estudiantes nuevos)

## Uso

1. Coloque su archivo Excel con los datos en la carpeta `data/` con el nombre `database.xlsx`, o especifique una ruta diferente usando el argumento `--input`.

2. Ejecute el programa:
```bash
python main.py
```

O especifique rutas personalizadas:
```bash
python main.py --input ruta/a/su/archivo.xlsx --output ruta/para/reportes/
```

## Resultados

El programa generará:
- Un archivo Excel por cada periodo analizado en el directorio de salida (por defecto: `reports/`)
- Cada archivo incluye:
  - Lista de estudiantes nuevos del periodo
  - Seguimiento de su continuidad en periodos posteriores
  - Estadísticas de retención
  - Formato condicional para facilitar la visualización

## Periodos Analizados

El software analiza la continuidad de estudiantes en los siguientes periodos:

- 1391 → 1394, 1395
- 1392 → 1394, 1395
- 1393 → 1394, 1395
- 1394 → 1701, 1702, 1703
- 1395 → 1701, 2033, 1703
- 1701 → 1704, 1705
- 1702 → 1704, 1705
- 1703 → 1704, 1705
- 1704 → 2031, 2032, 2033
- 1705 → 2031, 2032, 2033
- 2031 → 2034, 2035
- 2032 → 2034, 2035
- 2033 → 2034, 2035

## Características

- Análisis automático de múltiples periodos
- Seguimiento individual de estudiantes
- Cálculo de tasas de retención
- Reportes Excel con formato condicional
- Resumen de resultados en consola
- Manejo de errores robusto
- Logging detallado

## Estructura del Proyecto

```
proyecto/
├── README.md           # Este archivo
├── requirements.txt    # Dependencias del proyecto
├── config.py          # Configuraciones y mapeo de periodos
├── main.py           # Punto de entrada del programa
├── retention_analysis.py  # Lógica de análisis
└── excel_writer.py    # Generación de reportes Excel
```

## Notas Adicionales

- El análisis es case-insensitive para la condición "nuevo"
- Los reportes incluyen formato condicional (verde para continuidad, rojo para discontinuidad)
- Se generan estadísticas detalladas por periodo
- El programa incluye logging detallado para facilitar la depuración
