import argparse
import os
import logging
from retention_analysis import load_data, analyze_retention
from excel_writer import write_report
from config import DATA_FILE, OUTPUT_DIR

def parse_args():
    """
    Parse command line arguments for the application.
    """
    parser = argparse.ArgumentParser(
        description="Software de Análisis de Retención de Estudiantes",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "--input",
        type=str,
        default=DATA_FILE,
        help="Ruta del archivo Excel con la base de datos (default: data/database.xlsx)"
    )
    parser.add_argument(
        "--output",
        type=str,
        default=OUTPUT_DIR,
        help="Directorio de salida para los reportes (default: reports/)"
    )
    return parser.parse_args()

def setup_logging():
    """
    Configure logging settings for the application.
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

def print_summary(retention_results):
    """
    Print a summary of the analysis results.
    """
    print("\n=== Resumen del Análisis de Retención ===")
    print("-" * 50)
    
    for period, df in retention_results.items():
        if df.empty:
            continue
            
        print(f"\nPeriodo Inicial: {period}")
        print(f"Total de estudiantes nuevos: {len(df)}")
        
        # Calculate and display retention rates for each subsequent period
        continuity_cols = [col for col in df.columns if col.startswith('Continuidad en')]
        for col in continuity_cols:
            retention_rate = (df[col].sum() / len(df)) * 100
            next_period = col.replace('Continuidad en ', '')
            print(f"Tasa de retención en {next_period}: {retention_rate:.2f}%")
        
        print("-" * 30)

def main():
    """
    Main function that orchestrates the retention analysis process.
    """
    # Set up logging
    setup_logging()
    
    # Parse command line arguments
    args = parse_args()
    
    try:
        # Display startup message
        logging.info("Iniciando análisis de retención de estudiantes...")
        
        # Load the data
        logging.info(f"Cargando datos desde: {args.input}")
        df = load_data(args.input)
        logging.info(f"Datos cargados exitosamente. Total de registros: {len(df)}")
        
        # Analyze retention for all periods
        logging.info("Iniciando análisis de retención...")
        retention_results = analyze_retention(df)
        
        if not retention_results:
            logging.warning("No se encontraron estudiantes nuevos en ningún periodo.")
            return
        
        # Create output directory if it doesn't exist
        os.makedirs(args.output, exist_ok=True)
        
        # Write individual reports for each period
        for period, result_df in retention_results.items():
            if not result_df.empty:
                output_file = os.path.join(args.output, f"retencion_periodo_{period}.xlsx")
                write_report(result_df, output_file)
        
        # Print summary of results
        print_summary(retention_results)
        
        logging.info(f"Análisis completado. Los reportes se han guardado en: {args.output}")
        
        # Consolidar reportes en un único archivo
        logging.info("Consolidando reportes en un único archivo...")
        from consolidate_reports import consolidate_retention_reports
        consolidate_retention_reports()
        
    except FileNotFoundError as e:
        logging.error(f"Error: No se encontró el archivo de entrada. {str(e)}")
        return 1
    except Exception as e:
        logging.error(f"Error durante la ejecución: {str(e)}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())
