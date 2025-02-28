# XLSXWriter

## Overview

XLSXWriter is a PHP library built on top of [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) that allows you to easily convert data into an Excel (.xlsx) file. With its modular design, XLSXWriter provides a clean API for generating Excel files with customizable features such as headers, rows, borders, font styles, background colors, and more.

> **Note:** This library requires PHP 8.3 or later.

## Features

- **Modular Architecture:**  
  Separate classes handle building the spreadsheet, applying styles, and saving the file.
- **Customizable Styling:**  
  Easily customize fonts, colors, borders, and cell formats using PhpSpreadsheet’s style API.
- **Robust Excel Generation:**  
  Leverages PhpSpreadsheet to create real Excel files (.xlsx) with support for advanced features.
- **Easy Integration:**  
  Suitable for plain PHP projects or Laravel applications.

## Installation

Install XLSXWriter via Composer. This command will also install PhpSpreadsheet as a dependency:

```bash
composer require franco28dev/xlsxwriter
```

## Using Examples

### Basic Example

> - Below is a simple example that creates an Excel file with headers and data rows, applying basic styling (such as bold headers with a yellow background and thin borders):

```php
require 'vendor/autoload.php';

use XLSXWriter\ExcelWriter;

$excelWriter = new ExcelWriter();

// Set headers (automatically styled)
$excelWriter->setHeaders(['Name', 'Age', 'Email']);

// Add data rows
$excelWriter->addRow(['John Doe', 30, 'john@example.com']);
$excelWriter->addRow(['Jane Doe', 25, 'jane@example.com']);

// Save the Excel file to disk
$filename = 'output.xlsx';
if ($excelWriter->write($filename)) {
    echo "Excel file created successfully at {$filename}.";
} else {
    echo "Error creating Excel file.";
}
```

## Advanced Usage in a Laravel Controller

> - Here’s an example of how to integrate XLSXWriter in a Laravel controller to generate and download an Excel file:

```php
namespace App\Http\Controllers;

use Illuminate\Http\Request;
use XLSXWriter\ExcelWriter;

class ExcelController extends Controller
{
    public function downloadExcel()
    {
        // Initialize XLSXWriter
        $excelWriter = new ExcelWriter();
        $excelWriter->setHeaders(['Name', 'Age', 'Email'])
                    ->addRow(['John Doe', 30, 'john@example.com'])
                    ->addRow(['Jane Doe', 25, 'jane@example.com']);

        // Write the file to disk
        $filename = 'download.xlsx';
        if ($excelWriter->write($filename)) {
            // Return file as a download and delete after sending
            return response()->download($filename)->deleteFileAfterSend(true);
        }

        return response("Error creating Excel file.", 500);
    }
}
```

## Positioning Your Table with Offsets

```php
use XLSXWriter\ExcelWriter;

$excelWriter = new ExcelWriter();

// Set the offset to (3, 3) => D4
$excelWriter->setOffset(3, 3);

// Define headers and rows
$excelWriter->setHeaders(['Fecha', 'Total']);
$excelWriter->addRow(['2023-01-01', '1000']);
$excelWriter->addRow(['2023-01-02', '1500']);

// Save the Excel file to disk
$filename = 'offset_example.xlsx';
if ($excelWriter->write($filename)) {
    echo "Excel file created successfully at {$filename}.";
} else {
    echo "Error creating Excel file.";
}
```

## Customizing Styles

```php
// Optional: If you implement a getter for the spreadsheet:
$spreadsheet = $excelWriter->getSpreadsheet();
$spreadsheet->getActiveSheet()->getStyle('A1:C100')->applyFromArray([
    'font' => [
        'name' => 'Calibri',
        'size' => 11
    ]
]);
```
