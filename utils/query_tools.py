import pandas as pd
from sqlalchemy import exc
from pathlib import Path
from IPython.display import display
import dtale
import re

# Directory for exporting reports
output_directory = Path("C:/Reports")
output_directory.mkdir(parents=True, exist_ok=True)


def truncate(value, width):
    """
    Truncate long string values for compact preview (used in older table views).
    """
    return str(value)[:width] + "..." if len(str(value)) > width else str(value)


def fix_duplicate_columns_in_query(query):
    """
    Automatically adds descriptive aliases to columns in SELECT to prevent duplicates.
    Example: dbo.Acc_AccountsKol.Name_ → Acc_AccountsKol_Name_
    """
    select_match = re.search(r"SELECT\s+(.*?)\s+FROM", query, re.IGNORECASE | re.DOTALL)
    if not select_match:
        print("❌ Could not parse SELECT block.")
        return query

    select_block = select_match.group(1)
    columns = [col.strip() for col in select_block.split(',')]

    seen = {}
    aliased_columns = []

    for col in columns:
        # Match dbo.Table.Column or Table.Column
        table_match = re.search(r"(?:dbo\.)?(\w+)\.(\w+)", col)
        if table_match:
            table = table_match.group(1)
            column = table_match.group(2)
            alias_base = f"{table}_{column}"
        else:
            alias_base = col.replace('.', '_')

        alias = alias_base
        if alias in seen:
            seen[alias] += 1
            alias = f"{alias}_{seen[alias]}"
        else:
            seen[alias] = 1

        aliased_columns.append(f"{col} AS {alias}")

    new_select = ",\n    ".join(aliased_columns)
    fixed_query = re.sub(
        r"SELECT\s+(.*?)\s+FROM",
        f"SELECT\n    {new_select}\nFROM",
        query,
        flags=re.IGNORECASE | re.DOTALL
    )
    return fixed_query


def fetch_and_display_queries(
    engine,
    queries,
    pnum=10,
    export_to_excel=False,
    export_to_csv=False,
    display_tail=False,
    use_dtale=False,
    auto_fix_duplicates=True,
    show_dataframe=True,
):
    """
    Run multiple SQL queries, auto-fix duplicate columns if needed,
    preview results, export to Excel/CSV, and show interactive display in Jupyter.

    Parameters:
    - engine: SQLAlchemy engine object.
    - queries: List of SQL query strings.
    - pnum: Number of rows to preview.
    - export_to_excel: Export all results to one Excel file (multi-sheet).
    - export_to_csv: Export each query result to a separate CSV file.
    - display_tail: Show last pnum rows instead of first.
    - use_dtale: Open D-Tale interactive UI per query.
    - auto_fix_duplicates: Automatically fix SELECT duplicate columns if needed.
    - show_dataframe: Display DataFrame interactively in Jupyter.
    """

    if not engine:
        print("❌ No database connection.")
        return

    if export_to_excel:
        file_path = output_directory / f"report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        writer = pd.ExcelWriter(file_path, engine="openpyxl")

    for i, query in enumerate(queries, 1):
        print(f"\n🟡 Executing query {i}:\n{query}")
        try:
            # Try reading original query
            try:
                df = pd.read_sql(query, engine)
            except Exception as e:
                if "Field names must be unique" in str(e) and auto_fix_duplicates:
                    print("⚠️ Duplicate columns found. Fixing query...")
                    query = fix_duplicate_columns_in_query(query)
                    df = pd.read_sql(query, engine)
                else:
                    raise e

            # Preview portion
            preview_df = df.tail(pnum) if display_tail else df.head(pnum)

            if show_dataframe:
                print(f"📋 Preview of query {i} result (top {pnum} rows):")
                display(preview_df)

            print(f"📊 Query {i} returned {len(df)} records.")

            # Open in D-Tale (optional)
            if use_dtale:
                print(f"🌐 Opening D-Tale UI for query {i}...")
                dtale.show(df, ignore_duplicate=True).open_browser()

            # Export to Excel (multi-sheet)
            if export_to_excel:
                df.to_excel(writer, sheet_name=f"Query_{i}", index=False)
                print(f"✅ Exported to Excel sheet Query_{i}")

            # Export to CSV
            if export_to_csv:
                csv_path = output_directory / f"query_{i}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv"
                df.to_csv(csv_path, index=False)
                print(f"✅ Exported to CSV: {csv_path}")

        except exc.SQLAlchemyError as db_error:
            print(f"❌ Database error during query {i}: {db_error}")
        except Exception as e:
            print(f"❌ Unexpected error during query {i}: {e}")

    if export_to_excel:
        writer.close()
        print(f"\n✅ All data exported to: {file_path}")
