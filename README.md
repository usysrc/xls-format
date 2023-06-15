# xls-format

(＾ ▽ ＾)／＼(＾ ▽ ＾)

xls-format is a command-line tool for formatting columns in an Excel file. It allows you to specify the column range and the desired format (text or number) for the columns, and applies the formatting to the specified range in the Excel file.

## Installation

To install xls-format, you need to have Go installed. Then, you can use the following command to install the tool:

```shell
go install github.com/usysrc/xls-format
```

## Usage

To format columns in an Excel file, use the following command format:

```shell
xls-format <file-path> -s <sheet-index> -b <start-column> -e <end-column> -t <format-type>
```

The available options are as follows:

<file-path>: Path to the Excel file.
-s, --sheet <sheet-index>: Index of the sheet (starting from 0).
-b, --start <start-column>: Starting column (e.g., A).
-e, --end <end-column>: Ending column (e.g., Z).
-t, --format <format-type>: Column format: text or number.

Example usage:

```shell
xls-format sheet.xlsx -s 0 -b A -e B --format number
```
