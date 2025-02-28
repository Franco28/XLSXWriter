<?php

declare(strict_types=1);

namespace XLSXWriter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

class SpreadsheetBuilder
{
    private Spreadsheet $spreadsheet;

    /**
     * currentRow will track where we are as we write rows.
     */
    private int $currentRow;

    /**
     * By default, we'll start at D4 => column = 4, row = 4
     */
    private int $startColumnIndex = 4;
    private int $startRow = 4;

    /**
     * Keep track of how many headers exist, if needed for additional logic.
     */
    private int $headersCount = 0;

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();

        // Initialize currentRow to startRow (D4 => row = 4)
        $this->currentRow = $this->startRow;
    }

    /**
     * Allows you to define the starting column and row for the table.
     * For example, setOffset(3, 3) => starts at D4.
     *
     * @param int $columnIndex 1 = A, 2 = B, 3 = C, 4 = D, etc.
     * @param int $rowIndex   1 = first row, 2 = second row, etc.
     */
    public function setOffset(int $columnIndex, int $rowIndex): void
    {
        $this->startColumnIndex = $columnIndex;
        $this->startRow = $rowIndex;
        // Update currentRow so that subsequent rows build from here
        $this->currentRow = $rowIndex;
    }

    /**
     * Sets the headers (first row of the table) with a custom style.
     *
     * @param array $headers
     * @param array $customStyle e.g. ['borders' => false] to remove borders
     */
    public function setHeaders(array $headers, array $customStyle = []): void
    {
        // 1) Keep track of how many headers there are
        $this->headersCount = count($headers);

        // 2) Start writing headers at the defined offset
        $columnIndex = $this->startColumnIndex;
        $sheet = $this->spreadsheet->getActiveSheet();

        // We'll use this to map each header to a position in 'columnWidths'
        $headerPosition = 0;

        // 3) Write each header cell
        foreach ($headers as $headerName => $type) {
            $colLetter = $this->columnIndexToLetter($columnIndex);

            // Write the header text (the array key)
            $sheet->setCellValue($colLetter . $this->currentRow, $headerName);

            // 3.1) If 'columnWidths' is provided, check if there's a width for this column
            if (isset($customStyle['columnWidths'][$headerPosition])) {
                $desiredWidth = $customStyle['columnWidths'][$headerPosition];

                if ($desiredWidth === 'auto') {
                    // Enable auto-sizing for this column
                    $sheet->getColumnDimension($colLetter)->setAutoSize(true);
                } else {
                    // Set a fixed width (e.g., 20, 30, etc.)
                    $sheet->getColumnDimension($colLetter)->setWidth($desiredWidth);
                }
            }

            $columnIndex++;
            $headerPosition++;
        }

        // 4) Calculate the header range for styling (e.g., D4:F4)
        $firstColLetter = $this->columnIndexToLetter($this->startColumnIndex);
        $lastColLetter  = $this->columnIndexToLetter($columnIndex - 1);
        $range          = "{$firstColLetter}{$this->currentRow}:{$lastColLetter}{$this->currentRow}";

        // 5) Define default style for the headers
        $defaultStyle = [
            'font' => [
                'bold' => true,
                'size' => 16,
                'name' => 'Arial'
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'argb' => 'FFFFCC00' // Default background color (yellow)
                ]
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000']
                ]
            ]
        ];

        // 6) Merge user-provided style with the default style
        $finalStyle = array_replace_recursive($defaultStyle, $customStyle);

        // 7) If the user explicitly sets 'borders' => false, remove borders entirely
        if (isset($customStyle['borders']) && $customStyle['borders'] === false) {
            unset($finalStyle['borders']);
        }

        // 8) Apply the merged style to the header range
        $sheet->getStyle($range)->applyFromArray($finalStyle);

        // 9) Move down one row for data
        $this->currentRow++;
    }

    /**
     * Adds a row of data to the sheet, allowing optional custom style/borders.
     *
     * @param array $row
     * @param array $customStyle e.g. ['borders' => false] to remove borders
     */
    public function addRow(array $row, array $customStyle = []): void
    {
        $columnIndex = $this->startColumnIndex;
        $sheet = $this->spreadsheet->getActiveSheet();

        // 1) Write each cell
        foreach ($row as $cell) {
            $colLetter = $this->columnIndexToLetter($columnIndex);
            $sheet->setCellValue($colLetter . $this->currentRow, $cell);
            $columnIndex++;
        }

        // 2) Calculate the range for the newly added row (e.g., D5:F5)
        $firstColLetter = $this->columnIndexToLetter($this->startColumnIndex);
        $lastColLetter  = $this->columnIndexToLetter($columnIndex - 1);
        $range          = "{$firstColLetter}{$this->currentRow}:{$lastColLetter}{$this->currentRow}";

        // 3) Define a default style (thin borders)
        $defaultStyle = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000']
                ]
            ]
        ];

        // 4) Merge user-provided style
        $finalStyle = array_replace_recursive($defaultStyle, $customStyle);

        // 5) If the user sets 'borders' => false, remove borders entirely
        if (isset($customStyle['borders']) && $customStyle['borders'] === false) {
            unset($finalStyle['borders']);
        }

        // 6) Apply the final style to the row
        $sheet->getStyle($range)->applyFromArray($finalStyle);

        // 7) Move to the next row
        $this->currentRow++;
    }

    /**
     * Optional method to apply a custom style to any defined range.
     *
     * @param string $range       e.g. "A1:C10"
     * @param array  $styleArray
     */
    public function applyStyleToRange(string $range, array $styleArray): void
    {
        $this->spreadsheet->getActiveSheet()->getStyle($range)->applyFromArray($styleArray);
    }

    /**
     * Returns the constructed Spreadsheet object.
     *
     * @return Spreadsheet
     */
    public function build(): Spreadsheet
    {
        return $this->spreadsheet;
    }

    /**
     * Converts a 1-based column index to its corresponding letter.
     * For example, 1 => A, 2 => B, 3 => C, 4 => D, 26 => Z, 27 => AA, etc.
     *
     * @param int $columnIndex
     * @return string
     */
    private function columnIndexToLetter(int $columnIndex): string
    {
        $letter = '';
        while ($columnIndex > 0) {
            $remainder = ($columnIndex - 1) % 26;
            $letter = chr(65 + $remainder) . $letter;
            $columnIndex = intdiv($columnIndex - $remainder - 1, 26);
        }
        return $letter;
    }
}
