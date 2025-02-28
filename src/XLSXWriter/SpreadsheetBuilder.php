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
     */
    public function setHeaders(array $headers): void
    {
        $this->headersCount = count($headers);

        // Start writing headers at the defined offset
        $columnIndex = $this->startColumnIndex;

        foreach ($headers as $header) {
            $colLetter = $this->columnIndexToLetter($columnIndex);
            $this->spreadsheet->getActiveSheet()->setCellValue($colLetter . $this->currentRow, $header);
            $columnIndex++;
        }

        // Calculate the header range for styling (e.g., D4:F4)
        $firstColLetter = $this->columnIndexToLetter($this->startColumnIndex);
        $lastColLetter  = $this->columnIndexToLetter($columnIndex - 1);
        $range          = "{$firstColLetter}{$this->currentRow}:{$lastColLetter}{$this->currentRow}";

        // Define a style for the headers
        $headerStyle = [
            'font' => [
                'bold' => true,
                'size' => 12,
                'name' => 'Arial'
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'argb' => 'FFFFCC00' // Yellow background
                ]
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000']
                ]
            ]
        ];

        // Apply the header style
        $this->spreadsheet->getActiveSheet()->getStyle($range)->applyFromArray($headerStyle);

        // Move down one row for data
        $this->currentRow++;
    }

    /**
     * Adds a row of data to the sheet with basic borders.
     *
     * @param array $row
     */
    public function addRow(array $row): void
    {
        $columnIndex = $this->startColumnIndex;

        foreach ($row as $cell) {
            $colLetter = $this->columnIndexToLetter($columnIndex);
            $this->spreadsheet->getActiveSheet()->setCellValue($colLetter . $this->currentRow, $cell);
            $columnIndex++;
        }

        // Calculate the range for the newly added row (e.g., D5:F5)
        $firstColLetter = $this->columnIndexToLetter($this->startColumnIndex);
        $lastColLetter  = $this->columnIndexToLetter($columnIndex - 1);
        $range          = "{$firstColLetter}{$this->currentRow}:{$lastColLetter}{$this->currentRow}";

        // Apply simple border styling
        $dataStyle = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000']
                ]
            ]
        ];
        $this->spreadsheet->getActiveSheet()->getStyle($range)->applyFromArray($dataStyle);

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
