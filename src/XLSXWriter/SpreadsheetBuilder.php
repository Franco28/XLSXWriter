<?php

declare(strict_types=1);

namespace XLSXWriter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

class SpreadsheetBuilder
{
    private Spreadsheet $spreadsheet;
    private int $currentRow = 1;
    private int $headersCount = 0;

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
    }

    /**
     * Agrega los encabezados a la primera fila y aplica un estilo personalizado.
     *
     * @param array $headers
     */
    public function setHeaders(array $headers): void
    {
        $this->headersCount = count($headers);
        $column = 'A';
        foreach ($headers as $header) {
            $this->spreadsheet->getActiveSheet()->setCellValue($column . $this->currentRow, $header);
            $column++;
        }
        // Calcula el rango del encabezado (por ejemplo, A1:C1)
        $lastColumn = chr(ord('A') + $this->headersCount - 1);
        $range = "A{$this->currentRow}:{$lastColumn}{$this->currentRow}";

        // Define un estilo para los encabezados
        $headerStyle = [
            'font' => [
                'bold' => true,
                'size' => 12,
                'name' => 'Arial'
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'argb' => 'FFFFCC00' // Color de fondo amarillo
                ]
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000']
                ]
            ]
        ];

        // Aplica el estilo al rango de encabezados
        $this->spreadsheet->getActiveSheet()->getStyle($range)->applyFromArray($headerStyle);

        $this->currentRow++;
    }

    /**
     * Agrega una fila de datos a la hoja y aplica bordes básicos a la fila.
     *
     * @param array $row
     */
    public function addRow(array $row): void
    {
        $column = 'A';
        foreach ($row as $cell) {
            $this->spreadsheet->getActiveSheet()->setCellValue($column . $this->currentRow, $cell);
            $column++;
        }

        // Calcula el rango de la fila agregada
        $lastColumn = chr(ord('A') + count($row) - 1);
        $range = "A{$this->currentRow}:{$lastColumn}{$this->currentRow}";

        // Aplica un estilo de borde simple a la fila de datos
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
     * Método opcional para aplicar un estilo a un rango definido.
     *
     * @param string $range
     * @param array $styleArray
     */
    public function applyStyleToRange(string $range, array $styleArray): void
    {
        $this->spreadsheet->getActiveSheet()->getStyle($range)->applyFromArray($styleArray);
    }

    /**
     * Retorna el objeto Spreadsheet construido.
     *
     * @return Spreadsheet
     */
    public function build(): Spreadsheet
    {
        return $this->spreadsheet;
    }
}
