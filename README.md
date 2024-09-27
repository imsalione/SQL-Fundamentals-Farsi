# SQL-Fundamentals-Farsi

### Project Introduction

This project is aimed at developing an automated reporting tool that facilitates accessing data from SQL Server databases, executing SQL queries, and exporting the resulting data into structured reports. The tool allows users to extract data and save it in both Excel and CSV formats, providing flexibility in data handling and storage. This functionality is particularly beneficial in data-driven environments where efficient reporting and data management are critical.

### General Workflow of the Project

1. **Database Connection**: 
   The first step in the process involves establishing a connection to the SQL Server database. This is done using the `SQLAlchemy` library, which provides a robust interface for interacting with relational databases. The connection can be made with or without login credentials, allowing flexibility in different authentication environments (SQL or Windows authentication).

2. **Executing SQL Queries**:
   After establishing a connection, users can execute their desired SQL queries to retrieve data from the database. These queries can include filtering, sorting, or aggregating data as per business needs. The tool then fetches the data into a Pandas DataFrame, which makes it easier to manipulate and display the results.

3. **Data Display**:
   The retrieved data is displayed in the console in a tabular format. Users can specify how many rows they want to view, or the default will show the top 10 rows of the result set.

4. **Exporting Data**:
   A core feature of the project is the ability to export data into Excel and CSV formats. Users can choose whether to export the data:
   - **To Excel**: Each query result is saved in a new sheet of an Excel workbook. The sheet names are automatically generated based on the current date and time, ensuring that no data is overwritten.
   - **To CSV**: If CSV export is selected, the data is saved in a separate CSV file with a unique name based on the current timestamp. This allows for easy sharing and further analysis using standard spreadsheet software.

### Libraries and Modules Used

1. **Pandas**:
   Pandas is a widely-used Python library for data manipulation and analysis. In this project, it is responsible for:
   - Fetching and storing SQL query results in DataFrame format.
   - Exporting data to Excel and CSV files.
   
2. **SQLAlchemy**:
   SQLAlchemy is a popular SQL toolkit and Object Relational Mapper (ORM) that facilitates communication with SQL databases. It allows:
   - Easy connection management with SQL Server.
   - Safe and efficient execution of SQL queries.

3. **Openpyxl**:
   Openpyxl is a Python library used to read and write Excel files. It enables:
   - Adding query results to new sheets within an existing Excel workbook.
   
4. **Pathlib**:
   Pathlib is part of the Python standard library used for handling file and directory paths in an object-oriented way. It ensures:
   - The proper creation of directories.
   - Safe file path manipulation for exporting reports.
   
5. **Tabulate** (optional):
   The Tabulate library is used to display DataFrame contents in a neatly formatted table within the console. This enhances readability and allows users to view data summaries directly.

### Key Features

- **Dynamic Data Export**: The ability to export SQL query results dynamically into either Excel or CSV formats, based on user preferences.
- **Multi-Sheet Excel Reporting**: Each query is stored in a new sheet within the same Excel file, ensuring that no previous data is overwritten.
- **Automated Filename Management**: Both Excel sheets and CSV files are given unique names using timestamps, allowing multiple reports to be saved without conflicts.
- **Error Handling**: The project incorporates robust error handling mechanisms to ensure smooth operation, including database connection errors, file writing errors, and SQL execution issues.

### Conclusion

This project offers a versatile and automated solution for generating database reports in structured formats. By leveraging popular Python libraries such as Pandas, SQLAlchemy, and Openpyxl, it simplifies the process of extracting, displaying, and exporting data from SQL Server databases. With the added ability to handle both Excel and CSV outputs, it provides a flexible tool for data analysis and reporting needs.
