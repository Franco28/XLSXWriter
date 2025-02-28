<?php

declare(strict_types=1);

namespace XLSXWriter;

class ExcelWriter
{
    private SpreadsheetBuilder $spreadsheetBuilder;
    private FileSaver $fileSaver;

    public function __construct()
    {
        $this->spreadsheetBuilder = new SpreadsheetBuilder();
        $this->fileSaver = new FileSaver();
    }

    /**
     * Sets how many columns and rows to skip before writing data.
     * For example, setOffset(3, 3) => start at D4.
     *
     * @param int $columns Number of columns to skip (1-based)
     * @param int $rows    Number of rows to skip (1-based)
     * @return $this
     */
    public function setOffset(int $columns, int $rows): self
    {
        $this->spreadsheetBuilder->setOffset($columns, $rows);
        return $this;
    }

    /**
     * Sets the headers (first row of the table).
     *
     * @param array $headers
     * @param array $customStyle e.g. ['borders' => false] to remove borders
     * @return $this
     */
    public function setHeaders(array $headers, array $customStyle = []): self
    {
        $this->spreadsheetBuilder->setHeaders($headers, $customStyle);
        return $this;
    }

    /**
     * Adds a row of data to the sheet.
     *
     * @param array $row
     * @param array $customStyle e.g. ['borders' => false] to remove borders
     * @return $this
     */
    public function addRow(array $row, array $customStyle = []): self
    {
        $this->spreadsheetBuilder->addRow($row);
        return $this;
    }

    /**
     * (Optional) Applies a custom style array to any cell range (e.g. "A1:C10").
     * Returns $this for chainability.
     *
     * @param string $range
     * @param array  $styleArray
     * @return $this
     */
    public function applyStyleToRange(string $range, array $styleArray): self
    {
        $this->spreadsheetBuilder->applyStyleToRange($range, $styleArray);
        return $this;
    }

    /**
     * Builds the Excel file and saves it to the specified path.
     *
     * @param string $filePath
     * @return bool
     */
    public function write(string $filePath): bool
    {
        $spreadsheet = $this->spreadsheetBuilder->build();
        return $this->fileSaver->save($spreadsheet, $filePath);
    }
}
