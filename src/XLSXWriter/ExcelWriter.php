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
     * Sets default header style (merged into the built-in header style).
     *
     * @param array $style
     * @return $this
     */
    public function setDefaultHeaderStyle(array $style): self
    {
        $this->spreadsheetBuilder->setDefaultHeaderStyle($style);
        return $this;
    }

    /**
     * Sets default row style (merged into the built-in row style).
     *
     * @param array $style
     * @return $this
     */
    public function setDefaultRowStyle(array $style): self
    {
        $this->spreadsheetBuilder->setDefaultRowStyle($style);
        return $this;
    }

    /**
     * Sets default column widths for headers.
     *
     * @param array $widths
     * @return $this
     */
    public function setColumnWidths(array $widths): self
    {
        $this->spreadsheetBuilder->setColumnWidths($widths);
        return $this;
    }

    /**
     * Enables or disables basic formula injection protection.
     *
     * @param bool $enabled
     * @return $this
     */
    public function setSanitizeFormulas(bool $enabled): self
    {
        $this->spreadsheetBuilder->setSanitizeFormulas($enabled);
        return $this;
    }

    /**
     * Sets the active sheet title.
     *
     * @param string $title
     * @return $this
     */
    public function setSheetTitle(string $title): self
    {
        $this->spreadsheetBuilder->setSheetTitle($title);
        return $this;
    }

    /**
     * Freezes the sheet at the specified cell (e.g. "A2").
     *
     * @param string $cell
     * @return $this
     */
    public function freezePane(string $cell): self
    {
        $this->spreadsheetBuilder->freezePane($cell);
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
        $this->spreadsheetBuilder->addRow($row, $customStyle);
        return $this;
    }

    /**
     * Adds multiple rows to the sheet.
     *
     * @param array $rows
     * @param array $customStyle
     * @return $this
     */
    public function addRows(array $rows, array $customStyle = []): self
    {
        $this->spreadsheetBuilder->addRows($rows, $customStyle);
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

    /**
     * Returns the last error message from write(), if any.
     *
     * @return string|null
     */
    public function getLastError(): ?string
    {
        return $this->fileSaver->getLastError();
    }
}
