import json
from pathlib import Path
from utils.db_connection import create_connection


def load_db_config(config_path="config.json"):
    """
    Load database connection configuration from a JSON file.
    """
    config_file = Path(config_path)
    if not config_file.exists():
        raise FileNotFoundError(f"❌ Config file not found: {config_path}")

    with open(config_file, 'r', encoding='utf-8') as file:
        config = json.load(file)
    return config


def initialize_database_connection(config_path="config.json", show_sample=False):
    """
    Create and test database connection using config file.

    Parameters:
    - config_path: str, path to config.json
    - show_sample: bool, if True runs and displays a test query

    Returns:
    - db: create_connection instance
    - engine: SQLAlchemy engine
    """
    config = load_db_config(config_path)

    db = create_connection(
        server=config.get("server"),
        database=config.get("database"),
        username=config.get("username"),
        password=config.get("password")
    )

    engine = db.connect()

    if db.check_connection():
        if show_sample:
            query = "SELECT TOP 10 * FROM dbo.Acc_AccountsKol"
            df = db.run_query(query)
            if df is not None:
                print("✅ Sample query result:")
                try:
                    from IPython.display import display
                    display(df)
                except ImportError:
                    print(df)
        return db, engine
    else:
        print("❌ Database connection failed.")
        return None, None
