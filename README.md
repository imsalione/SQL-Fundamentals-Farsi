# SQL-Fundamentals-Farsi
***

### Project Introduction

This project is designed to create a tool for accessing SQL Server databases, executing various queries, and extracting data. The main objective is to automate the generation of comprehensive reports from database data and save them in Excel files. This tool is especially useful in environments where regular, structured reporting is required.

### General Workflow of the Project

1. **Database Connection**: First, a connection to the SQL Server database is established using server details, database name, and optional login credentials. This connection is managed via the SQLAlchemy library, which provides a powerful interface for interacting with SQL databases in Python.

2. **Executing SQL Queries**: After establishing the connection, the user can execute their desired queries to retrieve data. These queries can involve selecting, filtering, and sorting data based on the needs of the user.

3. **Data Display**: The retrieved data is displayed in a tabular format in the console. The number of rows to be displayed can be adjusted according to the user’s preference.

4. **Saving Output to Excel Files**: A key feature of this project is the ability to save the extracted data into Excel files. Each dataset returned from the queries is saved in a new sheet within the same Excel file to prevent data overlap and ensure the reports are stored systematically.

### Libraries and Modules Used

- **Pandas**: Used to read and manage data in DataFrame format, as well as to export the data into Excel files.
- **SQLAlchemy**: Manages database connections and query execution with the SQL Server.
- **Openpyxl**: Facilitates working with Excel files, including creating and editing sheets within existing files.
- **Pathlib**: Helps in managing file paths and ensuring files are correctly saved in the intended directory.
- **Tabulate**: Provides a clean and organized display of data in a tabular format in the console.

### Conclusion

This project serves as an efficient tool for users who need to generate and manage reports from SQL Server databases. By automating data extraction, processing, and saving the results into Excel files, the tool saves time and enhances the accuracy of data management.
