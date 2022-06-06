# Upload To SQL Database

[Elliot](https://github.com/lukasheinz92/elliot/blob/main/README.md#elliot)

Instructions:

1. Before using this tool, check if:
    - you are able to connect to the destination database
    - you have appropriate access rights in the destination database
    - you have proper connection driver installed on your computer (working with ADODB)

2. Required and optional parameters:
    - Source Table ... range of cells in an Excel sheet containing data you want to upload into <Target Table>
    - Target Table ... name of a table located in the database where you want to insert your data
    - Connection String ... connection string of a database where you want to insert your data
    - Query Before ... (optional) query which is performed before the data is uploaded to SQL database
    - Query After ... (optional) query which is performed after the data is uploaded to SQL database

3. Usefull information:
    - visit <https://connectionstrings.com> for more information about connection strings
    - ADODB connection is used to connect to destination database
    - data is uploaded to database based on matching column names, not column order
    - column order may differ between Source Table and Target Table in a database
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
