# ZCL_EXCEL_HANDLER

## Overview

ZCL_EXCEL_HANDLER is a general Excel upload/download handler class for SAP ABAP systems. It supports both XLSX and CSV file formats for uploading and downloading data between SAP and PC or application server.

## Features

- Upload and download CSV files
- Upload and download XLSX files
- Support for both PC and application server file handling
- Customizable CSV column separator
- Header line handling
- Automatic data type conversion and validation

## Main Methods

### Constructor

- `constructor`: Initializes the class with an optional CSV column separator

### CSV Operations

- `upload_csv`: Uploads a CSV file to an internal table
- `download_csv`: Downloads an internal table to a CSV file

### XLSX Operations

- `upload_xlsx`: Uploads an XLSX file to an internal table
- `download_xlsx`: Downloads an internal table to an XLSX file

## Usage

1. Create an instance of the class:

   ```abap
   DATA(lo_excel_handler) = NEW zcl_excel_handler( ).
   ```

2. Upload a CSV file:

   ```abap
   lo_excel_handler->upload_csv(
   EXPORTING
    iv_file_path = 'path/to/file.csv'
    iv_server    = abap_false
    iv_hdr_lines = 1
   IMPORTING
    ir_table     = REF #( your_internal_table )
   ).
   ```

3. Download a CSV file:

   ```abap
   DATA(lv_bytes_written) = lo_excel_handler->download_csv(
   EXPORTING
    iv_file_path = 'path/to/output.csv'
    iv_server    = abap_false
    iv_header    = abap_true
    ir_table     = REF #( your_internal_table )
   ).
   ```

4. Upload an XLSX file:

   ```abap
   lo_excel_handler->upload_xlsx(
   EXPORTING
    iv_file_path = 'path/to/file.xlsx'
    iv_server    = abap_false
    iv_hdr_lines = 1
   IMPORTING
    ir_table     = REF #( your_internal_table )
   ).
   ```

5. Download an XLSX file:

   ```abap
   DATA(lv_bytes_written) = lo_excel_handler->download_xlsx(
   EXPORTING
    iv_file_path = 'path/to/output.xlsx'
    iv_server    = abap_false
    ir_table     = REF #( your_internal_table )
   ).
   ```

### Notes

- All methods support upload and download on both PC (foreground) and application server (foreground/background)
- The class handles data type conversions and validations automatically
- Error handling is implemented using the ZCX_EXCEL_HANDLER exception class

This class provides a convenient and flexible way to handle Excel file operations in SAP ABAP systems, supporting both CSV and XLSX formats with various options for file location and data handling.

### Examples

Test folder contains an example program **zzexcel_test** for testing the class.
