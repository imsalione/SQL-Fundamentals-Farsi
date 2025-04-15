# utils/db_connection.py

import logging
import urllib
from sqlalchemy import create_engine
from sqlalchemy.exc import SQLAlchemyError

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def create_connection(server, database, username=None, password=None):
    """
    Create a connection to the SQL Server database using SQLAlchemy.

    Parameters:
    - server: str, the name of the SQL Server
    - database: str, the name of the database
    - username: str, optional, SQL Server username (for SQL authentication)
    - password: str, optional, SQL Server password (for SQL authentication)

    Returns:
    - engine: sqlalchemy.engine.base.Engine, the SQLAlchemy engine object or None if the connection fails
    """
    try:
        if username and password:
            odbc_str = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={server};"
                f"DATABASE={database};"
                f"UID={username};"
                f"PWD={password};"
                f"TrustServerCertificate=yes;"
            )
        else:
            odbc_str = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={server};"
                f"DATABASE={database};"
                f"Trusted_Connection=yes;"
                f"TrustServerCertificate=yes;"
            )

        odbc_encoded = urllib.parse.quote_plus(odbc_str)
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={odbc_encoded}")

        # Test connection
        with engine.connect() as conn:
            logging.info("Connection successful!")

        return engine

    except SQLAlchemyError as e:
        logging.error("Connection failed: %s", e)
        return None
