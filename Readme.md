# Excel Sheet Parsing and Generating Report

This tool parses an Excel sheet and generates reports based on the data. It supports filtering by class number and exporting reports in JSON or Markdown format.

## Usage

To run the tool, use the following command:

```bash
go run main.go
```

You can customize the behavior using the following flags:

- **`--class int`**: If specified, generates the report only for the given class number.
- **`--export string`**: Specifies the export type. Options are `json` for JSON reports or `md` for Markdown reports.
- **`--filename string`**: Path to the Excel file to read data from. Defaults to `"./CSF111_202425_01_GradeBook_stripped.xlsx"`.
- **`--sheet string`**: Name of the sheet in the Excel file. Defaults to `"CSF111_202425_01_GradeBook"`.

### Example Commands

- **Default Run**: 
  ```bash
  go run main.go
  ```

- **Specify File and Sheet**:
  ```bash
  go run main.go --filename=yourfile.xlsx --sheet=yourSheetName
  ```

- **Filter by Class and Export as JSON**:
  ```bash
  go run main.go --class=123 --export=json
  ```

- **Export as Markdown**:
  ```bash
  go run main.go --export=md
  ```

## Features

- **Data Validation**: Validates data before generating reports.
- **Flexible Reporting**: Supports JSON and Markdown report formats.
- **Class Filtering**: Allows filtering reports by class number.

## Requirements

- Go installed on your system.
- `github.com/mattn/go-sqlite3` and `github.com/xuri/excelize/v2` packages installed.
