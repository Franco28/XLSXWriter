<?php

declare(strict_types=1);

namespace XLSXWriter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class FileSaver
{
    private ?string $lastError = null;

    /**
     * Guarda el objeto Spreadsheet en la ruta indicada.
     *
     * @param Spreadsheet $spreadsheet
     * @param string $filePath
     * @return bool
     */
    public function save(Spreadsheet $spreadsheet, string $filePath): bool
    {
        $this->lastError = null;

        $filePath = trim($filePath);
        if ($filePath === '') {
            $this->lastError = 'File path cannot be empty.';
            return false;
        }

        if (strpos($filePath, "\0") !== false) {
            $this->lastError = 'File path contains invalid characters.';
            return false;
        }

        if (is_dir($filePath)) {
            $this->lastError = 'File path points to a directory.';
            return false;
        }

        $directory = dirname($filePath);
        if ($directory === '.' || $directory === '') {
            $directory = getcwd() ?: '';
        }

        if ($directory !== '' && !is_dir($directory)) {
            $this->lastError = 'Target directory does not exist.';
            return false;
        }

        if ($directory !== '' && !is_writable($directory)) {
            $this->lastError = 'Target directory is not writable.';
            return false;
        }

        try {
            $writer = new Xlsx($spreadsheet);
            $writer->save($filePath);
            return true;
        } catch (\Throwable $e) {
            $this->lastError = $e->getMessage();
            return false;
        }
    }

    /**
     * Returns the last error message from save(), if any.
     *
     * @return string|null
     */
    public function getLastError(): ?string
    {
        return $this->lastError;
    }
}
