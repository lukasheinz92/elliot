# SQL Editor

[Elliot](https://github.com/lukasheinz92/elliot/blob/main/README.md#elliot)

Instructions:
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
