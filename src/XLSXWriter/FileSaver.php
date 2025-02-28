<?php

declare(strict_types=1);

namespace XLSXWriter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class FileSaver
{
    /**
     * Guarda el objeto Spreadsheet en la ruta indicada.
     *
     * @param Spreadsheet $spreadsheet
     * @param string $filePath
     * @return bool
     */
    public function save(Spreadsheet $spreadsheet, string $filePath): bool
    {
        try {
            $writer = new Xlsx($spreadsheet);
            $writer->save($filePath);
            return true;
        } catch (\Exception $e) {
            // Aquí podrías loggear el error para depuración
            return false;
        }
    }
}
