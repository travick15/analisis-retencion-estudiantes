import pandas as pd
import os
from config import (
    DATA_FILE,
    COLUMN_PERIODO,
    COLUMN_DOCUMENTO,
    COLUMN_CONDICION,
    CONDICION_NUEVO,
    RETENTION_MAPPING
)
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)

def load_data(file_path=DATA_FILE):
    """
    Loads Excel data and verifies required columns exist.
    Returns the DataFrame.
    """
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except FileNotFoundError:
        logging.error(f"El archivo {file_path} no se encontró.")
        raise
    except Exception as e:
        logging.error(f"Error al cargar el archivo: {e}")
        raise

    # Check required columns
    for col in [COLUMN_PERIODO, COLUMN_DOCUMENTO, COLUMN_CONDICION]:
        if col not in df.columns:
            error_msg = f"Columna requerida '{col}' no encontrada en el Excel."
            logging.error(error_msg)
            raise ValueError(error_msg)
    return df

def filter_new_students(df, period):
    """
    Filters the DataFrame to obtain students from the given period with condition 'nuevo'.
    """
    # Use case-insensitive matching for the condition
    filtro = (df[COLUMN_PERIODO].astype(str) == str(period)) & \
             (df[COLUMN_CONDICION].str.lower() == CONDICION_NUEVO.lower())
    return df.loc[filtro].copy()

def check_student_continuity(df, documento, period):
    """
    Checks if a given student (by 'documento') appears in the specified period.
    Returns True if continuity is found, else False.
    """
    cond = (df[COLUMN_PERIODO].astype(str) == str(period)) & \
           (df[COLUMN_DOCUMENTO] == documento)
    return not df.loc[cond].empty

def analyze_retention(df):
    """
    Analyzes retention for each period according to RETENTION_MAPPING.
    Returns a dictionary where key = initial period and value = result DataFrame.
    """
    results = {}
    
    for period, next_periods in RETENTION_MAPPING.items():
        logging.info(f"Procesando periodo inicial: {period}")
        nuevos_df = filter_new_students(df, period)
        
        if nuevos_df.empty:
            logging.warning(f"No se encontraron estudiantes nuevos en el periodo {period}.")
            continue

        # Create a list to hold dict records for report
        records = []
        for index, row in nuevos_df.iterrows():
            doc = row[COLUMN_DOCUMENTO]
            record = {
                COLUMN_DOCUMENTO: doc,
                COLUMN_PERIODO: period,
                "Condición": row[COLUMN_CONDICION]
            }
            
            # Check in each subsequent period whether the student appears
            for next_p in next_periods:
                record[f"Continuidad en {next_p}"] = check_student_continuity(df, doc, next_p)
            
            # Add additional statistics
            continuity_count = sum(1 for next_p in next_periods 
                                 if check_student_continuity(df, doc, next_p))
            record["Total Periodos con Continuidad"] = continuity_count
            record["Porcentaje de Continuidad"] = (continuity_count / len(next_periods)) * 100
            
            records.append(record)
        
        # Create DataFrame with the records
        result_df = pd.DataFrame(records)
        
        # Add summary statistics
        if not result_df.empty:
            total_students = len(result_df)
            for next_p in next_periods:
                col_name = f"Continuidad en {next_p}"
                retention_rate = (result_df[col_name].sum() / total_students) * 100
                logging.info(f"Tasa de retención para {period} -> {next_p}: {retention_rate:.2f}%")
        
        results[period] = result_df

    return results
