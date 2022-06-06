# Elliot

*TODO license agreement*

Elliot is an Excel addin designed especially for analysts to provide some advanced functions. *TODO license agreement*. It consists of two components:

1. [Tools available via Ribbon](https://github.com/lukasheinz92/elliot/edit/main/README.md#1-tools)
2. [Functions](https://github.com/lukasheinz92/elliot/edit/main/README.md#2-functions)


## 1. Tools

Following chapters correspond to groups in the Excel Ribbon.

### 1.1 Format and Functions

By selecting cells you can format/transform values:

- Upper Case
- Lower Case
- Clear Blanks - in some cases empty cells act like a cells with a value, you can clear them by this option
- Number - switch number format with / without "1000 separator" (space), no decimal digits
- Date - format number as a date

### 1.2 Advanced Filter

Advanced filter allows you to filter multiple values at once with any separator (new line, comma, space etc.). You can easily copy-pase values you want to filter. For more information follow these [instructions](https://github.com/lukasheinz92/elliot/new/main#advanced-filter).

### 1.3 SQL

There are two tools - SQL Editor and Upload To SQL Database.

### 1.3.1 SQL Editor

SQL Editor allows you to use standard SQL query to work with your sheets. For more information follow these [instructions](https://github.com/lukasheinz92/elliot/new/main#sql-editor).

### 1.3.2 Upload to SQL Database

If you need to quickly upload your data from a sheet to your SQL database, you can use this tool. It is suitable only for small datasets due to slow performance. This tool is inspired by [https://www.excel-sql-server.com/excel-sql-server-import-export-using-vba.htm](https://www.excel-sql-server.com/excel-sql-server-import-export-using-vba.htm). For more information follow these [instructions](https://github.com/lukasheinz92/elliot/new/main#upload-to-sql-database).

### 1.4 Workbook

Shows name, path and full name of the workbook with option to copy this information to you clipboard.

### 1.5 Sheet Tools

You can duplicate, delete, move or copy your sheet(s). By using the Advanced option you can also hide, unhide, protect, unprotect or move your sheets in your workbook.

Status assigned to each sheet shows if sheet is visible (V) or hidden (H), if it is protected (P) or unprotected (U).

### 1.6 Go to Sheet

Buttons to navigate sheets in your workbook.

### 1.7 Information

Info button with link to Github.


## 2. Functions

### 2.1 RegExExtract(text, pattern)

Extracts content from a text string matching regular expression.
Use function: RegExExtract_help(n) with argument n=1 or n=2 to get web links with instructions how to use regular expressions in Excel
VBA reference: "Microsoft VBScript Regular Expressions 5.5"
  
  
### 2.2 RegExExtract_help(n)

Help function to RegExExtract() function. Returns web links with instructions how to use regular expressions in Excel.

### 2.3 ExtractNumbers(text)

Returns all digits from a text string.
  
### 2.4 ExtractCharacters(text)

Returns all characters (excluding digits) from a text string.
  
### 2.5 LTrim(text)

Removes spaces from the left of a text string.

### 2.6 RTrim(text)

Removes spaces from the right of a text string.
  
### 2.7 LCut(FromCell, n)

Removes n characters from the left of a text string.
  
### 2.8 RCut(FromCell, n)

Removes n characters from the right of a text string.
  
### 2.9 Quarter(FromDate)

Returns the quarter of a date (Q1 - Q4).
