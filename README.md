# SQL-Fundamentals-Farsi
***

<section style="direction: ltr; text-align: justify; margin: 30px;">

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

***

<section style="direction: rtl; text-align: justify; margin: 30px;">

### مقدمه پروژه

این پروژه با هدف توسعه یک ابزار گزارش‌دهی خودکار طراحی شده است که دسترسی به داده‌های پایگاه داده SQL Server، اجرای کوئری‌های SQL و ذخیره داده‌های به‌دست‌آمده به گزارش‌های ساختاریافته را فراهم می‌کند. این ابزار به کاربران اجازه می‌دهد داده‌ها را استخراج کرده و در هر دو فرمت Excel و CSV ذخیره کنند، که انعطاف‌پذیری در مدیریت و ذخیره‌سازی داده‌ها را ارائه می‌دهد. این قابلیت به‌ویژه در محیط‌های مبتنی بر داده که گزارش‌دهی و مدیریت کارآمد داده‌ها اهمیت زیادی دارند، بسیار مفید است.

### عملکرد کلی پروژه

1. **اتصال به پایگاه داده**: 
   اولین گام در این فرایند، ایجاد یک اتصال به پایگاه داده SQL Server است. این کار با استفاده از کتابخانه `SQLAlchemy` انجام می‌شود که یک رابط قدرتمند برای تعامل با پایگاه‌های داده رابطه‌ای فراهم می‌کند. اتصال می‌تواند با یا بدون نام کاربری و رمز عبور انجام شود و این امکان را فراهم می‌آورد که در محیط‌های مختلف احراز هویت (SQL یا Windows) استفاده شود.

2. **اجرای کوئری‌های SQL**:
   پس از ایجاد اتصال، کاربران می‌توانند کوئری‌های SQL دلخواه خود را اجرا کرده و داده‌های مورد نیاز را از پایگاه داده بازیابی کنند. این کوئری‌ها می‌توانند شامل فیلتر، مرتب‌سازی یا تجمیع داده‌ها باشند. سپس ابزار داده‌ها را به یک DataFrame از کتابخانه Pandas منتقل می‌کند که امکان مدیریت و نمایش راحت‌تر نتایج را فراهم می‌کند.

3. **نمایش داده‌ها**:
   داده‌های بازیابی‌شده به‌صورت جدولی در کنسول نمایش داده می‌شوند. کاربران می‌توانند تعداد ردیف‌هایی که می‌خواهند مشاهده کنند را مشخص کنند، یا به‌طور پیش‌فرض ۱۰ ردیف اول نمایش داده می‌شود.

4. **ذخیره داده‌ها**:
   ویژگی اصلی این پروژه، توانایی ذخیره داده‌ها به فرمت‌های Excel و CSV است. کاربران می‌توانند انتخاب کنند که داده‌ها به‌صورت زیر ذخیره شوند:
   - **در Excel**: هر نتیجه کوئری در یک شیت جدید از یک فایل Excel ذخیره می‌شود. نام شیت‌ها به‌صورت خودکار و بر اساس تاریخ و زمان فعلی ایجاد می‌شوند تا از بازنویسی داده‌ها جلوگیری شود.
   - **در CSV**: اگر گزینه CSV انتخاب شود، داده‌ها در یک فایل CSV جداگانه با نام منحصربه‌فرد (بر اساس زمان فعلی) ذخیره می‌شوند که امکان به اشتراک‌گذاری و تحلیل بیشتر با نرم‌افزارهای استاندارد صفحه‌گسترده را فراهم می‌کند.

### کتابخانه‌ها و ماژول‌های مورد استفاده

1. **Pandas**:
   Pandas یک کتابخانه بسیار پرکاربرد در پایتون برای تحلیل و مدیریت داده‌ها است. در این پروژه، مسئول:
   - بارگذاری و ذخیره نتایج کوئری SQL در قالب DataFrame.
   - ذخیره داده‌ها به فایل‌های Excel و CSV.
   
2. **SQLAlchemy**:
   SQLAlchemy یک ابزار SQL و Mapper شیء رابطه‌ای (ORM) محبوب است که ارتباط با پایگاه‌های داده SQL را تسهیل می‌کند. این ابزار امکان:
   - مدیریت آسان اتصال به SQL Server.
   - اجرای ایمن و کارآمد کوئری‌های SQL را فراهم می‌کند.

3. **Openpyxl**:
   Openpyxl یک کتابخانه پایتون برای خواندن و نوشتن فایل‌های Excel است که در این پروژه برای:
   - افزودن نتایج کوئری به شیت‌های جدید در یک فایل Excel موجود به کار می‌رود.
   
4. **Pathlib**:
   Pathlib بخشی از کتابخانه استاندارد پایتون برای مدیریت مسیرهای فایل و دایرکتوری به شیوه‌ای شیءگرا است. این کتابخانه تضمین می‌کند که:
   - دایرکتوری‌های مورد نیاز به‌درستی ایجاد شوند.
   - مدیریت ایمن مسیرهای فایل برای ذخیره گزارش‌ها انجام شود.
   
5. **Tabulate** (اختیاری):
   کتابخانه Tabulate برای نمایش محتوای DataFrame در قالب یک جدول مرتب‌شده در کنسول استفاده می‌شود. این کار خوانایی نتایج را بهبود می‌بخشد و به کاربران امکان می‌دهد خلاصه‌ای از داده‌ها را به‌راحتی مشاهده کنند.

### ویژگی‌های کلیدی

- **ذخیره سازی پویا**: امکان ذخیره سازی نتایج کوئری SQL به‌صورت پویا در فرمت‌های Excel یا CSV، بسته به نیاز کاربر.
- **گزارش‌دهی چندشیتی در Excel**: هر کوئری در یک شیت جدید از یک فایل Excel ذخیره می‌شود و از بازنویسی داده‌ها جلوگیری می‌شود.
- **مدیریت خودکار نام فایل‌ها**: شیت‌های Excel و فایل‌های CSV نام‌های منحصربه‌فردی بر اساس زمان فعلی دریافت می‌کنند که به کاربران اجازه می‌دهد چندین گزارش را بدون تداخل ذخیره کنند.
- **مدیریت خطا**: پروژه شامل مکانیسم‌های مدیریت خطا برای تضمین عملکرد صحیح است که شامل مدیریت خطاهای اتصال به پایگاه داده، خطاهای نوشتن فایل و اجرای کوئری‌های SQL می‌باشد.

### نتیجه‌گیری

این پروژه یک راهکار همه‌کاره و خودکار برای تولید گزارش‌های پایگاه داده در قالب‌های ساختاریافته ارائه می‌دهد. با استفاده از کتابخانه‌های پرکاربردی مانند Pandas، SQLAlchemy و Openpyxl، این ابزار فرایند استخراج، نمایش و ذخیره‌سازی داده از پایگاه داده SQL Server را ساده می‌کند. با داشتن قابلیت ذخیره‌سازی خروجی به‌صورت Excel و CSV، این ابزار یک ابزار انعطاف‌پذیر برای نیازهای تحلیل و گزارش‌دهی داده‌ها فراهم می‌کند.

</section>

***
