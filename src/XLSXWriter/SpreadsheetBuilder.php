<?php

declare(strict_types=1);

namespace XLSXWriter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

class SpreadsheetBuilder
{
    private Spreadsheet $spreadsheet;
    private array $defaultHeaderStyle;
    private array $defaultRowStyle;
    private array $columnWidths = [];
    private bool $sanitizeFormulas = false;

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

        $this->defaultHeaderStyle = [
            'font' => [
                'bold' => true,
                'size' => 15,
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

        $this->defaultRowStyle = [
            'font' => [
                'bold' => false,
                'size' => 12,
                'name' => 'Arial'
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'argb' => 'FFFFFFFF' // Default background color (white)
                ]
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000']
                ]
            ]
        ];

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
        $this->startColumnIndex = max(1, $columnIndex);
        $this->startRow = max(1, $rowIndex);
        // Update currentRow so that subsequent rows build from here
        $this->currentRow = $this->startRow;
    }

    /**
     * Sets default header style (merged into the built-in header style).
     *
     * @param array $style
     */
    public function setDefaultHeaderStyle(array $style): void
    {
        $this->defaultHeaderStyle = $this->mergeStyles(
            $this->defaultHeaderStyle,
            $this->filterStyleOverrides($style)
        );
    }

    /**
     * Sets default row style (merged into the built-in row style).
     *
     * @param array $style
     */
    public function setDefaultRowStyle(array $style): void
    {
        $this->defaultRowStyle = $this->mergeStyles(
            $this->defaultRowStyle,
            $this->filterStyleOverrides($style)
        );
    }

    /**
     * Sets default column widths to use in headers if not provided per call.
     *
     * @param array $widths
     */
    public function setColumnWidths(array $widths): void
    {
        $this->columnWidths = $widths;
    }

    /**
     * Enables or disables basic formula injection protection.
     *
     * @param bool $enabled
     */
    public function setSanitizeFormulas(bool $enabled): void
    {
        $this->sanitizeFormulas = $enabled;
    }

    /**
     * Sets the active sheet title (sanitized for Excel limits).
     *
     * @param string $title
     */
    public function setSheetTitle(string $title): void
    {
        $cleanTitle = $this->sanitizeSheetTitle($title);
        if ($cleanTitle === '') {
            return;
        }

        $this->spreadsheet->getActiveSheet()->setTitle($cleanTitle);
    }

    /**
     * Freezes the sheet at the specified cell (e.g. "A2").
     *
     * @param string $cell
     */
    public function freezePane(string $cell): void
    {
        $cell = trim($cell);
        if ($cell === '') {
            return;
        }

        $this->spreadsheet->getActiveSheet()->freezePane($cell);
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

        if ($this->headersCount === 0) {
            return;
        }

        // 2) Start writing headers at the defined offset
        $columnIndex = $this->startColumnIndex;
        $sheet = $this->spreadsheet->getActiveSheet();

        // We'll use this to map each header to a position in 'columnWidths'
        $headerPosition = 0;
        $columnWidths = $customStyle['columnWidths'] ?? $this->columnWidths;

        // 3) Write each header cell
        foreach ($headers as $headerName => $type) {
            $colLetter = $this->columnIndexToLetter($columnIndex);

            // Support both indexed and associative header arrays.
            $headerText = is_int($headerName) ? (string)$type : (string)$headerName;
            $sheet->setCellValue(
                $colLetter . $this->currentRow,
                $this->sanitizeCellValue($headerText)
            );

            // 3.1) If 'columnWidths' is provided, check if there's a width for this column
            if (is_array($columnWidths) && isset($columnWidths[$headerPosition])) {
                $desiredWidth = $columnWidths[$headerPosition];

                if ($desiredWidth === 'auto') {
                    // Enable auto-sizing for this column
                    $sheet->getColumnDimension($colLetter)->setAutoSize(true);
                } elseif (is_numeric($desiredWidth) && (float)$desiredWidth > 0) {
                    // Set a fixed width (e.g., 20, 30, etc.)
                    $sheet->getColumnDimension($colLetter)->setWidth((float)$desiredWidth);
                }
            }

            $columnIndex++;
            $headerPosition++;
        }

        // 4) Calculate the header range for styling (e.g., D4:F4)
        $firstColLetter = $this->columnIndexToLetter($this->startColumnIndex);
        $lastColLetter  = $this->columnIndexToLetter($columnIndex - 1);
        $range          = "{$firstColLetter}{$this->currentRow}:{$lastColLetter}{$this->currentRow}";

        // 5) Merge user-provided style with the default style
        $styleOverrides = $this->filterStyleOverrides($customStyle);
        $finalStyle = $this->mergeStyles($this->defaultHeaderStyle, $styleOverrides);

        // 6) If the user explicitly sets 'borders' => false, remove borders entirely
        if (isset($customStyle['borders']) && $customStyle['borders'] === false) {
            unset($finalStyle['borders']);
        }

        // 7) Apply the merged style to the header range
        $sheet->getStyle($range)->applyFromArray($finalStyle);

        // 8) Move down one row for data
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
        if (count($row) === 0) {
            $this->currentRow++;
            return;
        }

        $columnIndex = $this->startColumnIndex;
        $sheet = $this->spreadsheet->getActiveSheet();

        // 1) Write each cell
        foreach ($row as $cell) {
            $colLetter = $this->columnIndexToLetter($columnIndex);
            $sheet->setCellValue(
                $colLetter . $this->currentRow,
                $this->sanitizeCellValue($cell)
            );
            $columnIndex++;
        }

        $cellsWritten = max(count($row), $this->headersCount);
        if ($cellsWritten === 0) {
            $this->currentRow++;
            return;
        }

        // 2) Calculate the range for the newly added row (e.g., D5:F5)
        $firstColLetter = $this->columnIndexToLetter($this->startColumnIndex);
        $lastColLetter  = $this->columnIndexToLetter(
            $this->startColumnIndex + $cellsWritten - 1
        );
        $range          = "{$firstColLetter}{$this->currentRow}:{$lastColLetter}{$this->currentRow}";

        // 3) Merge user-provided style
        $styleOverrides = $this->filterStyleOverrides($customStyle);
        $finalStyle = $this->mergeStyles($this->defaultRowStyle, $styleOverrides);

        // 4) If the user sets 'borders' => false, remove borders entirely
        if (isset($customStyle['borders']) && $customStyle['borders'] === false) {
            unset($finalStyle['borders']);
        }

        // 5) Apply the final style to the row
        $sheet->getStyle($range)->applyFromArray($finalStyle);

        // 6) Move to the next row
        $this->currentRow++;
    }

    /**
     * Adds multiple rows to the sheet.
     *
     * @param array $rows
     * @param array $customStyle
     */
    public function addRows(array $rows, array $customStyle = []): void
    {
        foreach ($rows as $row) {
            $rowData = is_array($row) ? $row : [$row];
            $this->addRow($rowData, $customStyle);
        }
    }

    /**
     * Optional method to apply a custom style to any defined range.
     *
     * @param string $range       e.g. "A1:C10"
     * @param array  $styleArray
     */
    public function applyStyleToRange(string $range, array $styleArray): void
    {
        $range = trim($range);
        if ($range === '') {
            return;
        }

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
        if ($columnIndex < 1) {
            return '';
        }

        $letter = '';
        while ($columnIndex > 0) {
            $remainder = ($columnIndex - 1) % 26;
            $letter = chr(65 + $remainder) . $letter;
            $columnIndex = intdiv($columnIndex - $remainder - 1, 26);
        }
        return $letter;
    }

    /**
     * Removes non-style keys before applying to PhpSpreadsheet.
     *
     * @param array $style
     * @return array
     */
    private function filterStyleOverrides(array $style): array
    {
        if (array_key_exists('columnWidths', $style)) {
            unset($style['columnWidths']);
        }

        return $style;
    }

    /**
     * Merges style overrides into base style.
     *
     * @param array $base
     * @param array $overrides
     * @return array
     */
    private function mergeStyles(array $base, array $overrides): array
    {
        return array_replace_recursive($base, $overrides);
    }

    /**
     * Prevents formula injection if enabled.
     *
     * @param mixed $value
     * @return mixed
     */
    private function sanitizeCellValue(mixed $value): mixed
    {
        if (!$this->sanitizeFormulas || !is_string($value) || $value === '') {
            return $value;
        }

        $trimmed = ltrim($value, " \t");
        if ($trimmed !== '' && in_array($trimmed[0], ['=', '+', '-', '@'], true)) {
            return "'" . $value;
        }

        return $value;
    }

    /**
     * Sanitizes a sheet title for Excel constraints.
     *
     * @param string $title
     * @return string
     */
    private function sanitizeSheetTitle(string $title): string
    {
        $title = trim($title);
        if ($title === '') {
            return '';
        }

        $title = preg_replace('/[\\[\\]\\:\\*\\?\\/\\\\]/', ' ', $title) ?? '';
        $title = trim($title);
        if ($title === '') {
            return '';
        }

        if (strlen($title) > 31) {
            $title = substr($title, 0, 31);
        }

        return $title;
    }
}
