# XLSXWriter

## Overview

`XLSXWriter` is a PHP/Laravel library that allows you to easily convert data to an Excel file. It provides a simple and convenient way to generate Excel files with customizable features such as borders, font styles, and column widths.

## Installation

To use `XLSXWriter`, you can install it via Composer:

```bash
    composer require franco28dev/xlsxwriter
```

## Usage

#### Setting Excel File Properties

```php
    $writer->setTitle($title);
    $writer->setSubject($subject);
    $writer->setAuthor($author);
    $writer->setCompany($company);
    $writer->setKeywords($keywords);
    $writer->setDescription($description);
    $writer->setTempDir($tempdir);
    $writer->setRightToLeft($isRightToLeft);
```

#### Writing to File

```php
    $writer->writeToFile($filename);
```

#### Writing Sheet Header

```php
    $writer->writeSheetHeader($sheetName, $headerTypes, $columnOptions);
```

#### Writing Sheet Row

```php
    $writer->writeSheetRow($sheetName, $row, $rowOptions);
```

#### Marking Merged Cell

```php
    $writer->markMergedCell($sheet_name, $start_cell_row, $start_cell_column, $end_cell_row, $end_cell_column);
```

#### Writing Entire Sheet

```php
    $writer->writeSheet($data, $sheet_name, $header_types);
```

## Example

```php
public function GenerateDataToExcel()
{
    // Start Buffer
    ob_start();

    // Add the first array to scope everything, the second array as header, third array as values for the row
    $dataCajaDiaria = [
        ['Date', 'Total Month Of July'],
        [
            date('Y-m-d H:i:s'),
            '$150',
        ],
    ];

    // Initialize the class
    $writer = new XLSXWriter();

    // Set the Author description for the excel file
    $writer->setAuthor('Franco28 Dev');

    // Set the title tab
    $writer->writeSheet("July sales");

    // The filename
    $filename = 'july_sales.xlsx';

    // Write the file
    $writer->writeToFile($filename);

    // Clean Buffer for the download
    ob_get_clean();

    // If you want you can check if the file was created or not
    if (file_exists($filename)) {

        // Generate the headers for the download,
        // after the download the file will be deleted
        header('Content-Description: File Transfer');
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header('Content-Disposition: attachment; filename="' . basename($filename) . '"');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Length: ' . filesize($filename));
        readfile($filename);
        unlink($filename);
    }

    // Return the code to whatever you want
    return view("sales.month");
}
```
