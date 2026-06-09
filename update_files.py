"""
Update_files
===================
Sistema para hacer un update a un archivo de excel

Autor:  Emir_B
Fecha:  2026-03-23
"""

# ---------------------------------------------------------------------------
# Importaciones estándar
# ---------------------------------------------------------------------------
import os
import sys
import logging
from pathlib import Path

# ---------------------------------------------------------------------------
# Importaciones de terceros (pip install ...)
# ---------------------------------------------------------------------------
import time
import win32com.client

# ---------------------------------------------------------------------------
# Importaciones locales
# ---------------------------------------------------------------------------
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Configuración de logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO, 
    format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
VERSION = "1.0.0"
REFRESH_WAIT_SECONDS = 20
QUERY_KEYWORD = "consulta"

# ---------------------------------------------------------------------------
# Variables del enviroment
# ---------------------------------------------------------------------------

env_path = Path(r"C:\Owncloud\Codigos\.env")
load_dotenv(dotenv_path=env_path)

try:
    FILE_PATH_INVENTORY = os.getenv("FILE_PATH_INVENTORY")
    OUTPUT_FILE_INVENTORY = os.getenv("OUTPUT_FILE_INVENTORY")
except ValueError as error_file:
    logger.info(f"No se encontraron las direcciones de archivos'{error_file}'")
    print("\n")

# ---------------------------------------------------------------------------
# Clases
# ---------------------------------------------------------------------------



# ---------------------------------------------------------------------------
# Funciones
# ---------------------------------------------------------------------------

def matar_excel_total() -> bool:
    """
    Cierra todos los procesos ligados a excel.exe,
    de esta manera nos aseguramos de que el archivo siempre este
    en una operacion limpia y nueva
    """
    logger.info("Buscando y cerrando procesos de Excel...")
    codigo_salida = os.system("taskkill /F /IM excel.exe 2>nul")

    if codigo_salida == 0:
        logger.info("✅ Se han eliminado todas las instancias de Excel.")
        print("\n")
    else:
        logger.info("ℹ️ No se encontraron procesos de Excel abiertos.")
        print("\n")

def word_in_string(word: str, string: str) -> bool:
    """Verifica si 'word' aparece en 'string' (insensible a mayúsculas)."""
    return word.lower() in string.lower()


def _refresh_connection(conn, wb) -> bool:
    """
    Refresca una conexión individual.
    Retorna True si fue exitoso, False en caso contrario.
    """
    if not word_in_string(QUERY_KEYWORD, conn.Name):
        logger.info(f"Omitiendo conexión no-ODBC: '{conn.Name}'")
        return True

    try:
        # BackgroundQuery=False garantiza que Refresh() sea SÍNCRONO
        # Con True, el código continúa antes de que termine la actualización
        conn.OLEDBConnection.BackgroundQuery = False
        conn.Refresh()
        logger.info(f"Conexión '{conn.Name}' actualizada correctamente.")
        return True

    except Exception as e:
        logger.error(f"Error al refrescar '{conn.Name}': {e}")
        try:
            wb.Save()
        except Exception as save_err:
            logger.warning(f"No se pudo guardar tras el error: {save_err}")
        return False


def refresh_all_connections(file_path: str, retry_wait: int = REFRESH_WAIT_SECONDS) -> bool:
    """
    Metodo para actualizar un archivo de excel con coneccion ODBC desde python
    """

    matar_excel_total()

    excel = None
    wb = None
    success = True

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(file_path)

            # ✅ Desactiva la verificación de niveles de privacidad de Power Query
            wb.Queries  # Asegura que el motor de PQ esté inicializado
            excel.AutomationSecurity = 1  # Permite macros/conexiones automáticas

            # Configura privacidad para ignorar el aviso de combinación
            logger.info(f"Se detectaron '{len(wb.Connections)}' conexiones")
            print("\n")
            for conn in wb.Connections:
                try:
                    conn.OLEDBConnection.BackgroundQuery = False
                    logger.info(f"BackgroundQuery desactivado para '{conn.Name}'.")
                except Exception:
                    pass  # Algunas conexiones no tienen OLEDBConnection

        except Exception as e:
            logger.error(f"No se pudo abrir el archivo '{file_path}': {e}")
            return False

        for conn in wb.Connections:
            _refresh_connection(conn, wb)
            time.sleep(retry_wait)

        wb.Save()
        logger.info("Archivo guardado correctamente.")
        print("\n")

    except Exception as e:
        logger.exception(f"Error inesperado: {e}")
        success = False

    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass

    return success
# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------
def main() -> int:
    """Función principal del script."""
    logger.info("Iniciando v%s", VERSION)

    #matar_excel_total()

    refresh_all_connections(FILE_PATH_INVENTORY)

    logger.info("Ejecución finalizada correctamente.")
    return 0


if __name__ == "__main__":
    sys.exit(main())