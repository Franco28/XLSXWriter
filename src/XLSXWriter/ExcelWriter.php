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
     * Establece los encabezados de la hoja Excel.
     *
     * @param array $headers
     * @return $this
     */
    public function setHeaders(array $headers): self
    {
        $this->spreadsheetBuilder->setHeaders($headers);
        return $this;
    }

    /**
     * Agrega una fila de datos a la hoja.
     *
     * @param array $row
     * @return $this
     */
    public function addRow(array $row): self
    {
        $this->spreadsheetBuilder->addRow($row);
        return $this;
    }

    /**
     * Construye el archivo Excel y lo guarda en la ruta especificada.
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
