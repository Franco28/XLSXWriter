# XLSXWriter

## Overview

XLSXWriter is a small PHP library built on top of [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet). It provides a simple, chainable API for generating `.xlsx` files with headers, rows, offsets, and optional styling.

> **Requirement:** PHP 8.4+.

## Features

- Clean API for building spreadsheets and saving them to disk.
- Header and row styling with PhpSpreadsheet-style arrays.
- Offset support to position tables anywhere on the sheet.
- Minimal surface area: `setHeaders`, `addRow`, `applyStyleToRange`, `write`.

## Installation

```bash
composer require xlsxwriter/excel
```

## Quick Start

```php
require 'vendor/autoload.php';

use XLSXWriter\ExcelWriter;

$writer = new ExcelWriter();

$writer->setHeaders(['Name', 'Age', 'Email'])
       ->addRow(['John Doe', 30, 'john@example.com'])
       ->addRow(['Jane Doe', 25, 'jane@example.com']);

if ($writer->write('output.xlsx')) {
    echo "Excel file created successfully.";
} else {
    echo "Error creating Excel file.";
}
```

## Header Styling

`setHeaders()` accepts a style array compatible with PhpSpreadsheet. You can also disable borders by passing `['borders' => false]`.

```php
use XLSXWriter\ExcelWriter;

$writer = new ExcelWriter();

$headerStyle = [
    'font' => [
        'size' => 16,
        'name' => 'Calibri',
        'bold' => true,
        'color' => ['argb' => 'FF003366'],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        'startColor' => ['argb' => 'FFF2F2F2'],
    ],
];

$writer->setHeaders(['Product', 'Qty', 'Price'], $headerStyle)
       ->addRow(['Keyboard', 2, 49.90])
       ->addRow(['Mouse', 1, 19.90])
       ->write('styled_headers.xlsx');
```

## Row Styling

`addRow()` also accepts a style array. This applies only to the row you are adding.

```php
use XLSXWriter\ExcelWriter;

$writer = new ExcelWriter();

$writer->setHeaders(['Date', 'Total']);

$writer->addRow(['2024-01-01', 1000], [
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        'startColor' => ['argb' => 'FFEAF7FF'],
    ],
]);

$writer->addRow(['2024-01-02', 1500], [
    'borders' => false,
]);

$writer->write('row_styles.xlsx');
```

## Offsets (Positioning)

Use `setOffset($columnIndex, $rowIndex)` to start the table at a specific cell.

```php
use XLSXWriter\ExcelWriter;

$writer = new ExcelWriter();

// Start at D4 (column 4, row 4)
$writer->setOffset(4, 4)
       ->setHeaders(['City', 'Population'])
       ->addRow(['Madrid', 3223000])
       ->addRow(['Barcelona', 1620000])
       ->write('offset_example.xlsx');
```

## Custom Range Styling

You can style any range after building your table:

```php
use XLSXWriter\ExcelWriter;

$writer = new ExcelWriter();

$writer->setHeaders(['A', 'B', 'C'])
       ->addRow([1, 2, 3])
       ->applyStyleToRange('A1:C2', [
           'font' => [
               'name' => 'Calibri',
               'size' => 11,
           ],
       ])
       ->write('range_style.xlsx');
```

## Advanced Formatting (Dates, Currency, Alignment)

PhpSpreadsheet lets you apply number formats and alignment. Use `applyStyleToRange()` after writing the data.

```php
use XLSXWriter\ExcelWriter;

$writer = new ExcelWriter();

$writer->setHeaders(['Date', 'Amount', 'Status'])
       ->addRow(['2024-01-01', 1234.5, 'Paid'])
       ->addRow(['2024-01-02', 987.65, 'Pending']);

$writer->applyStyleToRange('A2:A3', [
    'numberFormat' => ['formatCode' => 'yyyy-mm-dd'],
]);

$writer->applyStyleToRange('B2:B3', [
    'numberFormat' => ['formatCode' => '"$"#,##0.00'],
    'alignment' => ['horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT],
]);

$writer->applyStyleToRange('C2:C3', [
    'alignment' => ['horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER],
]);

$writer->write('formatted.xlsx');
```

## Laravel Example

```php
namespace App\Http\Controllers;

use XLSXWriter\ExcelWriter;

class ReportController extends Controller
{
    public function download()
    {
        $writer = new ExcelWriter();

        $writer->setHeaders(['Name', 'Score'])
               ->addRow(['Alice', 95])
               ->addRow(['Bob', 88]);

        $path = storage_path('app/reports/report.xlsx');

        if ($writer->write($path)) {
            return response()->download($path)->deleteFileAfterSend(true);
        }

        return response('Error creating Excel file.', 500);
    }
}
```

## API Summary

```php
ExcelWriter::setOffset(int $columns, int $rows): self
ExcelWriter::setHeaders(array $headers, array $customStyle = []): self
ExcelWriter::addRow(array $row, array $customStyle = []): self
ExcelWriter::applyStyleToRange(string $range, array $styleArray): self
ExcelWriter::write(string $filePath): bool
```

## Error Handling

- `write()` returns `true` on success and `false` on failure.
- If you need error details, wrap the call or update `FileSaver` to log exceptions.

## Limitations and Performance Notes

- This library writes in memory through PhpSpreadsheet; very large datasets will use significant RAM.
- Auto-sizing columns (if you enable it) can be slow on large sheets.
- Streaming writers are not included here; if you need huge exports, consider a streaming approach in PhpSpreadsheet.

## Compatibility

- PHP 8.4+
- PhpSpreadsheet ^1.28

## Testing

```bash
composer install
composer test
```

## Roadmap

- Multi-sheet support.
- Optional streaming writer for large exports.
- Helper methods for common formats (dates, currency, percentages).

## Contributing

Issues and PRs are welcome. Keep changes small and focused, and include a short description and tests when applicable.

## Changelog

See `composer.json` for the current version.

## Notes

- `setHeaders()` accepts both indexed arrays (`['Name', 'Email']`) and associative arrays (`['Name' => 'string']`).
- `write()` returns `true` on success and `false` on failure.
- Styling arrays follow the PhpSpreadsheet format. See their docs for full options.
