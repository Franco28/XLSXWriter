# XLSXWriter

## Overview

`XLSXWriter` is a PHP/Laravel library that allows you to easily convert data to an Excel file. It provides a simple and convenient way to generate Excel files with customizable features such as borders, font styles, and column widths.

## How it Works

To generate an Excel file, follow these steps:

1. **Start Buffer**: Begin output buffering using `ob_start()`.

2. **Prepare Data**: Organize your data in a multi-dimensional array. The first array is used to scope everything, the second array serves as the header, and the third array contains the values for each row.

3. **Initialize XLSXWriter**: Create an instance of the `XLSXWriter` class.

4. **Set Author and Title**: Use `setAuthor` to set the author description and `writeSheet` to define the title tab.

5. **Customize Appearance**:

   - **Enable Borders**: Utilize `setEnableBorders(true)` to enable borders for cells.
   - **Set Font Style**: Specify the font style using `setFont(['name' => 'Arial', 'size' => 12])`.

6. **Set Column Widths**: Adjust the column widths with `setColumnWidths([15, 20])`.

7. **Define Filename**: Provide a filename for the Excel file, such as `'july_sales.xlsx'`.

8. **Write to File**: Save the Excel file using `writeToFile($filename)`.

9. **Clean Buffer**: Clear the output buffer with `ob_get_clean()`.

10. **Download Excel File**: Check if the file was created, generate download headers, and delete the file after download.

11. **Return Code or View**: Return the desired code or view.

## Example

```php
public function GenerateDataToExcel()
{
    // Start Buffer
    ob_start();

    // Add the first array to scope everything, the second array as header, third array as values for the row
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

    // Enable borders for cells
    $writer->setEnableBorders(true);

    // Set font style
    $font = ['name' => 'Arial', 'size' => 12];
    $writer->setFont($font);

    // Set column widths
    $columnWidths = [15, 20]; // Sample widths for the two columns
    $writer->setColumnWidths($columnWidths);

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
