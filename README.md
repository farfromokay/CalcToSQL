# CalcToSQL v1.0

A LibreOffice Calc macro to generate MySQL `CREATE TABLE` and `INSERT INTO` statements from spreadsheet data.

## Description

CalcToSQL scans the sheets in your LibreOffice Calc document, interprets headers and data, and generates a SQL script file containing `CREATE TABLE` statements (including inferred data types, primary keys, unique keys, foreign keys, and indexes based on header hints) and `INSERT INTO` statements for the data in each sheet.

It allows for configuration via a dedicated sheet (`_CalcToSQL`) and uses simple hints in column headers to define database schema details.

## Features

* Generates `CREATE TABLE` statements from sheet data.
* Generates `INSERT INTO` statements for all data rows.
* Infers basic SQL data types (`INT`, `INT UNSIGNED`, `DOUBLE`, `DATETIME`, `VARCHAR(255)`) based on column content.
* Supports header hints for defining:
    * Primary Keys (`pk`)
    * Unique Keys (`uidx:<name>`)
    * Foreign Keys (`fk:referenced_table(referenced_column)`)
    * Indexes (`idx:<name>`)
    * Timestamp Columns (`tsc`): Sets type to `TIMESTAMP DEFAULT CURRENT_TIMESTAMP` and excludes the column from `INSERT` statements.
* Handles composite keys (multiple columns with the same `uidx` or `idx` name, or multiple `pk` hints).
* Automatically adds `AUTO_INCREMENT` to single-column integer primary keys.
* Allows skipping sheets or columns by prefixing their names with an underscore (`_`).
* Provides optional configuration via a `_CalcToSQL` sheet:
    * `DropIfExists` setting (TRUE/FALSE) to add `DROP TABLE IF EXISTS` statements.
    * Centralized `Hints` section to define hints without modifying sheet headers.
* Outputs a single `.sql` file compatible with MySQL.
* Includes `SET FOREIGN_KEY_CHECKS=0;` at the beginning and `SET FOREIGN_KEY_CHECKS=1;` at the end of the script.

## Requirements

* LibreOffice Calc (tested on recent versions, but should work on older versions supporting Basic macros)

## Installation / Usage

1.  **Open LibreOffice Calc.**
2.  Go to `Tools` -> `Macros` -> `Organize Macros` -> `LibreOffice Basic...`.
3.  In the "Macro from" dropdown, select your Calc document (or "My Macros" if you want it available globally).
4.  Click `New...`. Give the module a name (e.g., `CalcToSQL_Module`).
5.  Delete any existing template code in the editor window.
6.  Copy the entire code from the `CalcToSQL.bas` file.
7.  Paste the code into the LibreOffice Basic editor window.
8.  Save the macro (`File` -> `Save`).
9.  Close the Basic IDE.

**To Run the Macro:**

1.  Open the Calc document containing the sheets you want to convert.
2.  Ensure your sheets are formatted correctly (headers in Row 1, data starting in Row 2).
3.  Add any desired header hints (see below) or configure the `_CalcToSQL` sheet.
4.  Go to `Tools` -> `Macros` -> `Run Macro...`.
5.  Select your macro library (e.g., the name of your document or "My Macros"), then the module name (e.g., `CalcToSQL_Module`), and finally the `CalcToSQL` macro.
6.  Click `Run`.
7.  A "Save As" dialog will appear. Choose a location and filename for the output `.sql` file.
8.  A confirmation message box will appear upon completion, showing statistics.

**Example File:**

An example spreadsheet, `CalcToSQL_Example.ods`, is included in the repository. It demonstrates:
* A barebones sample data sheet named `Sheet1` with a single column header `Col1` and data values 1 through 5.
* A pre-configured `_CalcToSQL` sheet showing examples of the `Settings` and `Hints` sections.
You can use this file to test the macro and see the configuration options in action.

## Sheet Formatting

* **Sheet Names:** Each sheet name (that doesn't start with `_`) becomes a table name. Special characters in sheet names might cause issues; stick to alphanumeric names and underscores.
* **Header Row:** Row 1 **must** contain the column headers.
* **Data Rows:** Data should start from Row 2 downwards.
* **Skipping Sheets:** Any sheet whose name starts with an underscore (`_`) will be ignored (e.g., `_Notes`, `_CalcToSQL`).
* **Skipping Columns:** Any column whose header name in Row 1 starts with an underscore (`_`) will be ignored entirely (not included in `CREATE` or `INSERT`).

## Header Hints

Hints are added in square brackets `[]` at the end of the column header text in Row 1. Multiple hints are separated by commas. Hints are case-insensitive. Hints can also be defined in the `_CalcToSQL` sheet (see Configuration).

* `Column Name [pk]` - Marks this column as part of the primary key.
* `Column Name [uidx:index_name]` - Adds this column to a unique index named `index_name`. Multiple columns can share the same `uidx` name for a composite unique key.
* `Column Name [idx:index_name]` - Adds this column to a non-unique index named `index_name`. Multiple columns can share the same `idx` name for a composite index.
* `Column Name [fk:OtherTable(other_column)]` - Creates a foreign key constraint referencing `other_column` in `OtherTable`. Table and column names are case-sensitive as provided here.
* `Column Name [tsc]` - Sets the column type to `TIMESTAMP DEFAULT CURRENT_TIMESTAMP` and excludes it from `INSERT` statements. Data type inference is skipped for this column.

**Example Header:** `user_id [pk, fk:users(id), idx:user_lookup]`

## Configuration (`_CalcToSQL` Sheet)

You can optionally create a sheet named exactly `_CalcToSQL` to control settings and define hints centrally.

**1. Settings Section:**

Create a section like this (headers in Row 1, data starting Row 4):

| A        | B     |
| :------- | :---- |
| Settings |       |
|          |       |
| Name     | Value |
| dropifexists | TRUE |

* **`dropifexists`**: If set to `TRUE`, `Yes`, `1`, or `On` (case-insensitive), the script will add `DROP TABLE IF EXISTS table_name;` before each `CREATE TABLE` statement. Any other value (or if the setting/section is missing) defaults to `FALSE`.

**2. Hints Section:**

Create a section like this (headers in Row N, data starting Row N+3):

| A     | B        | C                                 |
| :---- | :------- | :-------------------------------- |
| Hints |          |                                   |
|       |          |                                   |
| Sheet | Column   | Hint                              |
| users | user_id  | pk, idx:user_id                   |
| posts | created_at | tsc                             |
| posts | user_id  | fk:users(id), idx:post_author     |
| posts | post_id  | pk                                |

* Hints defined here are *combined* with any hints found in the actual column header brackets `[]`.
* Duplicate hints resulting from the combination are ignored.
* This allows defining common constraints (like `tsc` for timestamp columns) without cluttering the main sheet headers.
* `Sheet` and `Column` names are case-insensitive when matching. The `Hint` string is processed as described in the "Header Hints" section.

## License

This project is licensed under the MIT License. See the license text included in the `CalcToSQL.bas` file.

## Author

* FarFromOkay (Replace with your name/handle if desired)
