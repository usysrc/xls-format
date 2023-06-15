# xls-format

xls-format is a command-line tool for formatting columns in an Excel file. It allows you to specify the column range and the desired format (text or number) for the columns, and applies the formatting to the specified range in the Excel file.

## Installation

To install xls-format, you need to have Go installed. Then, you can use the following command to install the tool:

```shell
go install github.com/usysrc/xls-format
```

## How to use

You can format some columns for example like this:

```shell
xls-format --file sheet.xlsx --start A --end B --format number
```
