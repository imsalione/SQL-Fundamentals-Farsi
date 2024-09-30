# SQL-Fundamentals-Farsi

### Project Overview (Updated)

#### 1. **Database Connection**
   The tool uses the powerful **SQLAlchemy** library to manage connections to the SQL Server database. The connection can be made with or without a username and password, allowing for flexible authentication in various environments (SQL or Windows authentication).

#### 2. **Executing SQL Queries**
   Users can execute their desired SQL queries. The retrieved data is directly loaded into a **DataFrame** from the **Pandas** library, enabling more straightforward and optimized handling of results. Additionally, the tool displays the number of records retrieved from each query, giving the user a clear view of the volume of data returned.

#### 3. **Data Display**
   The retrieved data is displayed in a table format in the console. For improved presentation and readability, the tool uses the **PrettyTable** library, which neatly organizes the tables. Users can specify the number of rows they wish to view, with a default setting of 10 rows if not specified.

#### 4. **Data Saving**
   Query results can be saved in both **Excel** and **CSV** formats. The data is automatically saved with unique names to avoid overwriting previous reports.

---

### Libraries and Modules Used (Updated)

#### 1. **Pandas**:
   - Loading and saving query results as DataFrames.
   - Saving data into Excel and CSV files.

#### 2. **SQLAlchemy**:
   - Managing connections to SQL Server and executing SQL queries.
   - Handling errors related to connections and query execution.

#### 3. **Openpyxl**:
   - Saving data into Excel files.
   - Managing multiple sheets within an Excel file.

#### 4. **Pathlib**:
   - Managing file paths and directories in an object-oriented way.
   - Ensuring that necessary directories are created for saving the output files.

#### 5. **PrettyTable**:
   - Displaying data in a neat, tabular format in the console.
   - Enhancing the readability and organization of data for the user.

---

### Key Features (Updated)

1. **Dynamic Saving**: Output data can be saved dynamically in Excel and CSV formats depending on the user's needs.
2. **File Naming Management**: Excel and CSV files are automatically named based on the current date and time.
3. **Error Management**: Error-handling mechanisms are implemented to manage issues related to connection and query execution.
4. **Multi-sheet Reporting in Excel**: Each query is saved in a new sheet within an Excel file, without overwriting previous sheets.
5. **Graphical Display of Results (Optional)**: The tool allows for the graphical representation of query results, such as charts.
6. **Record Count Display**: After executing each query, the number of retrieved records is displayed to the user.

---

### Conclusion

This project offers a versatile and automated solution for generating database reports in structured formats. By leveraging popular Python libraries such as Pandas, SQLAlchemy, and Openpyxl, it simplifies the process of extracting, displaying, and exporting data from SQL Server databases. With the added ability to handle both Excel and CSV outputs, it provides a flexible tool for data analysis and reporting needs.
