# Elliot

*TODO license agreement*

Elliot is an Excel addin designed especially for analysts to provide some advanced functions. *TODO license agreement*. It consists of two components:

1. Tools available via Ribbon
2. Functions

## 1. Tools

Following chapters correspond to groups in the Excel Ribbon.

### Format and Functions

By selecting cells you can format/transform values:

- Upper Case
- Lower Case
- Clear Blanks - in some cases empty cells act like a cells with a value, you can clear them by this option
- Number - switch number format with / without "1000 separator" (space), no decimal digits
- Date - format number as a date

### Advanced Filter

Advanced filter allows you to filter multiple values at once with any separator (new line, comma, space etc.). You can easily copy-pase values you want to filter. Follow these instructions:

1. Select any cell in a column of the table you want to filter.
2. Type or copy-paste values you want to filter into the text field (filter is not case sensitive).
3. Every value has to be on a new row (default) or separated by a delimiter. Choose your delimiter in the right section of the window. You can use preset delimiter or "Other" option with custom delimiter (custom delimiter may contain one or more characters, even numbers).
4. Empty rows at the bottom are ignored, else they are part of values to filter.
5. Spaces between single values and delimiters (even row delimiter) are trimmed.
6. Check "Keep Filter Open" if you want to keep the window open after filtering.

Error may occur primarily when:
- selected cell is out of the table
- invalid characters are entered into the Filter Criteria window
- you are trying to filter pivot table

### SQL

There are two tools - SQL Editor and Upload To SQL Database.

**SQL Editor**

SQL Editor allows you to use standard SQL query to work with your sheets. Follow these instructions:

1. You can use SELECT, UPDATE and INSERT INTO queries.
2. Use a sheet name as a table name in SQL query (very long sheet name may not work).
3. If a column or table name contains space/s, put them into [...] (square brackets).
4. No column in the query can have the exact same name as any sheet in the workbook.
5. Column and table aliases are allowed.
6. Query is not case sensitive except table names (sheet names).
7. MS Access query syntax (functions, INNER JOIN instead of JOIN only, etc.).
8. Keep only one table per sheet (preferred left top corner in A1).

Data Types:
- strings are defined by '...' (single quotes)
- for date use Datevalue(...) function; Datevalue(2018-09-26)
- NULL is allowed

You can use these wild cards (with LIKE):
- % (percentage) - various number of characters (including no character)
- _ (underscore) - exactly one character

SELECT / UPDATE / INSERT INTO:

a) SELECT:
- result of the SELECT query is inserted into new or existing sheet
- use SELECT Query Insert INTO field on the right of the SQL Editor
  to select existing sheet or to type a name of new sheet (will be added)
- if you keep it empty, you will be prompted to add new sheet

b) UPDATE:
- query updates a table on the same sheet

c) INSERT INTO:
- query inserts values into existing table
- you can use INSERT INTO Table1 SELECT * FROM Table2 or INSERT
  INTO Table1 Values(...)
- keep exactly one space between these two keywords (INSERT_INTO)

Error may be caused by:
- syntax error (typo)
- case sensitivity for table names
- using a column name with the exactly same name as any sheet name in the workbook
- multiple tables in one sheet

**Upload to SQL Database**

If you need to quickly upload your data from a sheet to your SQL database, you can use this tool. It is suitable only for small datasets due to slow performance. Please follow these instructions:

1. Before using this tool, check if:
    - you are able to connect to destination database
    - you have appropriate access rights in the database
    - you have proper connection driver installed on your computer (working with ADODB)

2. This tool has these required and optional parameters:
    - Source Table ... range of cells in an Excel sheet containing data you want to upload into <Target Table>
    - Target Table ... name of a table located in the database where you want to insert your data
    - Connection String ... connectiong string of a database where you want to insert your data
    - Query Before ... optional query which is performed before data are inserted into SQL database
    - Query After ... optional query which is performed after data are inserted into SQL database

3. Usefull information:
    - visit <https://connectionstrings.com> for more information about connection strings
    - ADODB connection is used to connect to destination database
    - data are inserted into database based on column names, not column order
    - column order may differ between <Source Table> and <Target Table> in a database
    - at least one pair of column names (Excel / database) must match
    - <Target Table> must exist (it is not created if does not exist)
    - use <Query Before> or <Query After> if you want to transform data, create or truncate tables etc.
    - you can enter multiple queries into <Query Before> or <Query After> - use ; (semicolon) as delimiter
    - optional queries must be supported by the vendor of a destination database
    - transaction is used during the export process
    - if all steps pass then transaction is committed
    - if at least one step fail then transaction is rolled back
    - the process of data load has following steps:
      1. Connection to database (using ADODB)
      2. Transaction Begin
      3. Run <Query Before> (if used)
      4. Insert data from <Source Table> into <Target Table>
      5. Run <Query After> (if used)
      6. Transaction Commit
      7. Close connection

4. Major error cases:
    - missing paramaters
    - connection failure
    - incorrect connection string
    - connection driver missing (not installed)
    - <Query Before> or <Query After> failure
    - <Source Table> does not contain header
    - <Target Table> does not exist
    - no column names match (between <Source Table> and <Target Tabe>) 

    When an error occurs, window with error message appears. No changes in the database are made because transaction is used keep data consistent (if transaction began then it is rolled back). If connection to database was already establisted then it was also closed.

This tool is inspired by:
https://www.excel-sql-server.com/excel-sql-server-import-export-using-vba.htm


### Workbook

Shows name, path and full name of the workbook with option to copy this information to you clipboard.

### Sheet Tools

TODO

### Go to Sheet

Buttons to navigate sheets in your workbook.

### Information

Info button with link to Github.


## 2. Functions
