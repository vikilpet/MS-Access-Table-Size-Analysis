# MS Access Table Size Analysis 

## How it works
Collect all non-system tables in database.

Export each table to a temporary database and compare size before and after.

Show the table with the collected information and delete the temporary database.

## Usage
Copy this Sub to a global module and run it with F5.

Don\'t forget to delete the temporary table (*Const StTable*).

## Prerequisites
Tested with MS Access 2010 and .mdb file.

If your file contains a table with multiple fields, you will get error 3838.

## \*\*\*
Feel free to ask questions or correct my spelling mistakes.
