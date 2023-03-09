# XLSXWriter

Convert data to excel file PHP / LARAVEL

## How it works?

```php
    public function GenerateDataToExcel()
    {
        // Start Buffer
        ob_start();

        // Add the first array to scope everything, the second array as header, thirth array as values for the row
        $dataCajaDiaria = [
            ['Date', 'Total Month Of July'],
            [
                date('Y-m-d H:i:s'),
                '$150',
            ],
        ];

        // Initialize the class
        $writer = new XLSXWriter();

        // Set the Author description for the excel file
        $writer->setAuthor('Franco28 Dev');

        // Set the title tab
        $writer->writeSheet("July sales");

        // The filename
        $filename = 'july_sales.xlsx';

        // Write the file
        $writer->writeToFile($filename);

        // Clean Buffer for the download
        ob_get_clean();

        // If you want you can check if the file was created or not
        if (file_exists($filename)) {

            // Generate the headers for the download,
            // after the download the file will be deleted
            header('Content-Description: File Transfer');
            header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            header('Content-Disposition: attachment; filename="' . basename($filename) . '"');
            header('Expires: 0');
            header('Cache-Control: must-revalidate');
            header('Pragma: public');
            header('Content-Length: ' . filesize($filename));
            readfile($filename);
            unlink($filename);
        }

        // Return the code to whatever you want
        return view("sales.month");
    }
```
