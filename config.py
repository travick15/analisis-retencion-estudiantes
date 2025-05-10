# Mapping for retention analysis.
# Each key is the initial period where "nuevos" (new students) are identified.
# The value is a list of subsequent Periodos where continuity is expected.
RETENTION_MAPPING = {
    "1391": ["1394", "1395"],
    "1392": ["1394", "1395"],
    "1393": ["1394", "1395"],
    "1394": ["1701", "1702", "1703"],
    "1395": ["1701", "2033", "1703"],
    "1701": ["1704", "1705"],
    "1702": ["1704", "1705"],
    "1703": ["1704", "1705"],
    "1704": ["2031", "2032", "2033"],
    "1705": ["2031", "2032", "2033"],
    "2031": ["2034", "2035"],
    "2032": ["2034", "2035"],
    "2033": ["2034", "2035"],
}

# File paths
DATA_FILE = "data/database.xlsx"  # Path to the Excel database file
OUTPUT_DIR = "reports/"           # Output directory for the retention analysis reports

# Required column names in the Excel (exact match expected)
COLUMN_PERIODO = "Periodo"       # Found in cell A1
COLUMN_DOCUMENTO = "Documento"   # Found in cell C1
COLUMN_CONDICION = "Condición"   # Found in cell I1

# Value to filter new students
CONDICION_NUEVO = "Nuevo"  # Case insensitive matching will be used in the code

# Logging configuration
LOGGING_LEVEL = "INFO"
