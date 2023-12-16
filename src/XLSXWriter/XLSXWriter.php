<?php

namespace XLSXWriter;

use ZipArchive;
use Exception;

class XLSXWriter
{
    //http://www.ecma-international.org/publications/standards/Ecma-376.htm
    //http://officeopenxml.com/SSstyles.php
    //------------------------------------------------------------------
    //http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
    const EXCEL_2007_MAX_ROW = 1048576;
    const EXCEL_2007_MAX_COL = 16384;
    //------------------------------------------------------------------
    protected $title;
    protected $subject;
    protected $author;
    protected $isRightToLeft;
    protected $company;
    protected $description;
    protected $tempdir;
    protected $keywords = array();

    protected $current_sheet;
    protected $sheets = array();
    protected $temp_files = array();
    protected $cell_styles = array();
    protected $number_formats = array();

    protected $columnWidths = [];
    protected $font = [];
    protected $enableBorders = false;

    /**
     * Set the title of the Excel file.
     *
     * @param string $title The title to set.
     */
    public function setTitle($title = '')
    {
        $this->title = $title;
    }

    /**
     * Set the subject of the Excel file.
     *
     * @param string $subject The subject to set.
     */
    public function setSubject($subject = '')
    {
        $this->subject = $subject;
    }

    /**
     * Set the author of the Excel file.
     *
     * @param string $author The author to set.
     */
    public function setAuthor($author = '')
    {
        $this->author = $author;
    }

    /**
     * Set the company associated with the Excel file.
     *
     * @param string $company The company to set.
     */
    public function setCompany($company = '')
    {
        $this->company = $company;
    }

    /**
     * Set keywords for the Excel file.
     *
     * @param string $keywords Keywords to set.
     */
    public function setKeywords($keywords = '')
    {
        $this->keywords = $keywords;
    }

    /**
     * Set the description of the Excel file.
     *
     * @param string $description The description to set.
     */
    public function setDescription($description = '')
    {
        $this->description = $description;
    }

    /**
     * Set the temporary directory for Excel file creation.
     *
     * @param string $tempdir The temporary directory path.
     */
    public function setTempDir($tempdir = '')
    {
        $this->tempdir = $tempdir;
    }

    /**
     * Set the right-to-left text direction for the Excel file.
     *
     * @param bool $isRightToLeft Whether to set right-to-left text direction.
     */
    public function setRightToLeft($isRightToLeft = false)
    {
        $this->isRightToLeft = $isRightToLeft;
    }

    /**
     * Enable or disable cell borders in the Excel file.
     *
     * @param bool $enableBorders Whether to enable or disable cell borders.
     */
    public function setEnableBorders($enableBorders)
    {
        $this->enableBorders = $enableBorders;
    }

    /**
     * Set column widths for the Excel file.
     *
     * @param array $widths An associative array of column widths (e.g., ['A' => 50]).
     */
    public function setColumnWidths($widths = ['A' => 50])
    {
        $this->columnWidths = $widths;
    }

    /**
     * Set the font style for the Excel file.
     *
     * @param array $font An associative array specifying the font (e.g., ['name' => 'Arial', 'size' => 12]).
     */
    public function setFont($font = ['name' => 'Arial', 'size' => 12])
    {
        $this->font = $font;
    }

    /**
     * XLSXWriter constructor.
     *
     * Initializes the XLSXWriter class, sets default configurations, and performs essential checks.
     */
    public function __construct()
    {
        defined('ENT_XML1') or define('ENT_XML1', 16);
        date_default_timezone_get() or date_default_timezone_set('UTC');
        is_writeable($this->tempFilename()) or self::log("Warning: tempdir " . sys_get_temp_dir() . " not writeable, use ->setTempDir()");
        class_exists('ZipArchive') or self::log("Error: ZipArchive class does not exist");
        $this->addCellStyle($number_format = 'GENERAL', $style_string = null);
    }

    /**
     * XLSXWriter destructor.
     *
     * Cleans up temporary files created during the XLSXWriter instance lifecycle.
     * Deletes any temporary files associated with the instance to free up resources.
     */
    public function __destruct()
    {
        if (!empty($this->temp_files)) {
            foreach ($this->temp_files as $temp_file) {
                @unlink($temp_file);
            }
        }
    }

    protected function tempFilename()
    {
        $tempdir = !empty($this->tempdir) ? $this->tempdir : sys_get_temp_dir();
        $filename = tempnam($tempdir, "xlsx_writer_");
        if (!$filename) {
            // If you are seeing this error, it's possible you may have too many open
            // file handles. If you're creating a spreadsheet with many small inserts,
            // it is possible to exceed the default 1024 open file handles. Run 'ulimit -a'
            // and try increasing the 'open files' number with 'ulimit -n 8192'
            throw new \Exception("Unable to create tempfile - check file handle limits?");
        }
        $this->temp_files[] = $filename;
        return $filename;
    }

    public function writeToStdOut()
    {
        $temp_file = $this->tempFilename();
        self::writeToFile($temp_file);
        readfile($temp_file);
    }

    public function writeToString()
    {
        $temp_file = $this->tempFilename();
        self::writeToFile($temp_file);
        $string = file_get_contents($temp_file);
        return $string;
    }

    public function writeToFile($filename)
    {
        foreach ($this->sheets as $sheet_name => $sheet) {
            self::finalizeSheet($sheet_name);
        }

        if (file_exists($filename)) {
            if (!is_writable($filename)) {
                self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not writeable.");
                return;
            }
            @unlink($filename);
        }

        $zip = new ZipArchive();
        if (!$zip->open($filename, ZipArchive::CREATE)) {
            self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", unable to create zip.");
            return;
        }

        $this->addZipDir($zip, "docProps/", array("app.xml", "core.xml"), array($this->buildAppXML(), $this->buildCoreXML()));
        $this->addZipDir($zip, "_rels/", array(".rels"), array($this->buildRelationshipsXML()));
        $this->addZipDir($zip, "xl/worksheets/", $this->getSheetFileNames(), $this->getSheetFiles());
        $zip->addFromString("xl/workbook.xml", $this->buildWorkbookXML());
        $zip->addFile($this->writeStylesXML(), "xl/styles.xml");
        $zip->addFromString("[Content_Types].xml", $this->buildContentTypesXML());
        $this->addZipDir($zip, "xl/_rels/", array("workbook.xml.rels"), array($this->buildWorkbookRelsXML()));

        $zip->close();
    }

    protected function initializeSheet($sheet_name, $col_widths = [], $auto_filter = false, $freeze_rows = false, $freeze_columns = false)
    {
        if ($this->current_sheet == $sheet_name || isset($this->sheets[$sheet_name])) {
            return;
        }

        $sheet_filename = $this->tempFilename();
        $sheet_xmlname = 'sheet' . (count($this->sheets) + 1) . ".xml";
        $this->sheets[$sheet_name] = (object)[
            'filename' => $sheet_filename,
            'sheetname' => $sheet_name,
            'xmlname' => $sheet_xmlname,
            'row_count' => 0,
            'file_writer' => new XLSXWriter_BuffererWriter($sheet_filename),
            'columns' => [],
            'merge_cells' => [],
            'max_cell_tag_start' => 0,
            'max_cell_tag_end' => 0,
            'auto_filter' => $auto_filter,
            'freeze_rows' => $freeze_rows,
            'freeze_columns' => $freeze_columns,
            'finalized' => false,
        ];

        $sheet = &$this->sheets[$sheet_name];
        $rightToLeftValue = $this->isRightToLeft ? 'true' : 'false';
        $tabselected = count($this->sheets) == 1 ? 'true' : 'false';
        $max_cell = XLSXWriter::xlsCell(self::EXCEL_2007_MAX_ROW, self::EXCEL_2007_MAX_COL); // XFE1048577

        $sheet->file_writer->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $sheet->file_writer->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
        $sheet->file_writer->write('<sheetPr filterMode="false">');
        $sheet->file_writer->write('<pageSetUpPr fitToPage="false"/>');
        $sheet->file_writer->write('</sheetPr>');
        $sheet->max_cell_tag_start = $sheet->file_writer->ftell();
        $sheet->file_writer->write('<dimension ref="A1:' . $max_cell . '"/>');
        $sheet->max_cell_tag_end = $sheet->file_writer->ftell();
        $sheet->file_writer->write('<sheetViews>');
        $sheet->file_writer->write('<sheetView colorId="64" defaultGridColor="true" rightToLeft="' . $rightToLeftValue . '" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="' . $tabselected . '" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');

        if ($sheet->freeze_rows && $sheet->freeze_columns) {
            $this->writePaneAndSelection($sheet, $sheet->freeze_rows, $sheet->freeze_columns);
        } elseif ($sheet->freeze_rows) {
            $this->writePaneAndSelection($sheet, $sheet->freeze_rows);
        } elseif ($sheet->freeze_columns) {
            $this->writePaneAndSelection($sheet, null, $sheet->freeze_columns);
        } else {
            $sheet->file_writer->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
        }

        $sheet->file_writer->write('</sheetView>');
        $sheet->file_writer->write('</sheetViews>');
        $sheet->file_writer->write('<cols>');

        $i = 0;
        if (!empty($col_widths)) {
            foreach ($col_widths as $column_width) {
                $this->writeCol($sheet, $i + 1, $column_width);
                $i++;
            }
        }

        $this->writeCol($sheet, $i + 1, 11.5, false, true);
        $sheet->file_writer->write('</cols>');
        $sheet->file_writer->write('<sheetData>');
    }

    private function addCellStyle($number_format, $cell_style_string)
    {
        $number_format_idx = self::add_to_list_get_index($this->number_formats, $number_format);
        $lookup_string = $number_format_idx . ";" . $cell_style_string;
        $cell_style_idx = self::add_to_list_get_index($this->cell_styles, $lookup_string);
        return $cell_style_idx;
    }

    private function initializeColumnTypes($header_types)
    {
        $column_types = array();
        foreach ($header_types as $v) {
            $number_format = self::numberFormatStandardized($v);
            $number_format_type = self::determineNumberFormatType($number_format);
            $cell_style_idx = $this->addCellStyle($number_format, $style_string = null);
            $column_types[] = array(
                'number_format' => $number_format, //contains excel format like 'YYYY-MM-DD HH:MM:SS'
                'number_format_type' => $number_format_type, //contains friendly format like 'datetime'
                'default_cell_style' => $cell_style_idx,
            );
        }
        return $column_types;
    }

    public function writeSheetHeader($sheet_name, array $header_types, $col_options = null)
    {
        if (empty($sheet_name) || empty($header_types) || !empty($this->sheets[$sheet_name]))
            return;

        $suppress_row = isset($col_options['suppress_row']) ? intval($col_options['suppress_row']) : false;
        if (is_bool($col_options)) {
            self::log("Warning! passing $suppress_row=false|true to writeSheetHeader() is deprecated, this will be removed in a future version.");
            $suppress_row = intval($col_options);
        }
        $style = &$col_options;

        $col_widths = isset($col_options['widths']) ? (array)$col_options['widths'] : array();
        $auto_filter = isset($col_options['auto_filter']) ? intval($col_options['auto_filter']) : false;
        $freeze_rows = isset($col_options['freeze_rows']) ? intval($col_options['freeze_rows']) : false;
        $freeze_columns = isset($col_options['freeze_columns']) ? intval($col_options['freeze_columns']) : false;
        self::initializeSheet($sheet_name, $col_widths, $auto_filter, $freeze_rows, $freeze_columns);
        $sheet = &$this->sheets[$sheet_name];
        $sheet->columns = $this->initializeColumnTypes($header_types);
        if (!$suppress_row) {
            $header_row = array_keys($header_types);

            $sheet->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . (1) . '">');
            foreach ($header_row as $c => $v) {
                $cell_style_idx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle('GENERAL', json_encode(isset($style[0]) ? $style[$c] : $style));
                $this->writeCell($sheet->file_writer, 0, $c, $v, $number_format_type = 'n_string', $cell_style_idx);
            }
            $sheet->file_writer->write('</row>');
            $sheet->row_count++;
        }
        $this->current_sheet = $sheet_name;
    }

    public function writeSheetRow($sheet_name, array $row, $row_options = null)
    {
        if (empty($sheet_name))
            return;

        $this->initializeSheet($sheet_name);
        $sheet = &$this->sheets[$sheet_name];
        if (count($sheet->columns) < count($row)) {
            $default_column_types = $this->initializeColumnTypes(array_fill($from = 0, $until = count($row), 'GENERAL')); //will map to n_auto
            $sheet->columns = array_merge((array)$sheet->columns, $default_column_types);
        }

        if (!empty($row_options)) {
            $ht = isset($row_options['height']) ? floatval($row_options['height']) : 12.1;
            $customHt = isset($row_options['height']) ? true : false;
            $hidden = isset($row_options['hidden']) ? (bool)($row_options['hidden']) : false;
            $collapsed = isset($row_options['collapsed']) ? (bool)($row_options['collapsed']) : false;
            $sheet->file_writer->write('<row collapsed="' . ($collapsed ? 'true' : 'false') . '" customFormat="false" customHeight="' . ($customHt ? 'true' : 'false') . '" hidden="' . ($hidden ? 'true' : 'false') . '" ht="' . ($ht) . '" outlineLevel="0" r="' . ($sheet->row_count + 1) . '">');
        } else {
            $sheet->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($sheet->row_count + 1) . '">');
        }

        $style = &$row_options;
        $c = 0;
        foreach ($row as $v) {
            $number_format = $sheet->columns[$c]['number_format'];
            $number_format_type = $sheet->columns[$c]['number_format_type'];
            $cell_style_idx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle($number_format, json_encode(isset($style[0]) ? $style[$c] : $style));
            $this->writeCell($sheet->file_writer, $sheet->row_count, $c, $v, $number_format_type, $cell_style_idx);
            $c++;
        }
        $sheet->file_writer->write('</row>');
        $sheet->row_count++;
        $this->current_sheet = $sheet_name;
    }

    public function countSheetRows($sheet_name = '')
    {
        $sheet_name = $sheet_name ? $sheet_name : $this->current_sheet;
        return array_key_exists($sheet_name, $this->sheets) ? $this->sheets[$sheet_name]->row_count : 0;
    }

    protected function finalizeSheet($sheet_name)
    {
        if (empty($sheet_name) || $this->sheets[$sheet_name]->finalized)
            return;

        $sheet = &$this->sheets[$sheet_name];

        $sheet->file_writer->write('</sheetData>');

        if (!empty($sheet->merge_cells)) {
            $sheet->file_writer->write('<mergeCells>');
            foreach ($sheet->merge_cells as $range) {
                $sheet->file_writer->write('<mergeCell ref="' . $range . '"/>');
            }
            $sheet->file_writer->write('</mergeCells>');
        }

        $max_cell = self::xlsCell($sheet->row_count - 1, count($sheet->columns) - 1);

        if ($sheet->auto_filter) {
            $sheet->file_writer->write('<autoFilter ref="A1:' . $max_cell . '"/>');
        }

        $sheet->file_writer->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
        $sheet->file_writer->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
        $sheet->file_writer->write('<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
        $sheet->file_writer->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $sheet->file_writer->write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
        $sheet->file_writer->write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
        $sheet->file_writer->write('</headerFooter>');
        $sheet->file_writer->write('</worksheet>');

        $max_cell_tag = '<dimension ref="A1:' . $max_cell . '"/>';
        $padding_length = $sheet->max_cell_tag_end - $sheet->max_cell_tag_start - strlen($max_cell_tag);
        $sheet->file_writer->fseek($sheet->max_cell_tag_start);
        $sheet->file_writer->write($max_cell_tag . str_repeat(" ", $padding_length));
        $sheet->file_writer->close();
        $sheet->finalized = true;
    }

    public function markMergedCell($sheet_name, $start_cell_row, $start_cell_column, $end_cell_row, $end_cell_column)
    {
        if (empty($sheet_name) || $this->sheets[$sheet_name]->finalized)
            return;

        self::initializeSheet($sheet_name);
        $sheet = &$this->sheets[$sheet_name];

        $startCell = self::xlsCell($start_cell_row, $start_cell_column);
        $endCell = self::xlsCell($end_cell_row, $end_cell_column);
        $sheet->merge_cells[] = $startCell . ":" . $endCell;
    }

    public function writeSheet(array $data, $sheet_name = '', array $header_types = array())
    {
        $sheet_name = empty($sheet_name) ? 'Sheet1' : $sheet_name;
        $data = empty($data) ? array(array('')) : $data;
        if (!empty($header_types)) {
            $this->writeSheetHeader($sheet_name, $header_types);
        }
        foreach ($data as $i => $row) {
            $this->writeSheetRow($sheet_name, $row);
        }
        $this->finalizeSheet($sheet_name);
    }

    protected function writeCell(XLSXWriter_BuffererWriter &$file, $row_number, $column_number, $value, $num_format_type, $cell_style_idx)
    {
        $cell_name = self::xlsCell($row_number, $column_number);

        if (!is_scalar($value) || $value === '') { //objects, array, empty
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '"/>');
        } elseif (is_string($value) && $value[0] == '=') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="s"><f>' . self::xmlspecialchars($value) . '</f></c>');
        } elseif ($num_format_type == 'n_date') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . intval(self::convert_date_time($value)) . '</v></c>');
        } elseif ($num_format_type == 'n_datetime') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . self::convert_date_time($value) . '</v></c>');
        } elseif ($num_format_type == 'n_numeric') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . self::xmlspecialchars($value) . '</v></c>'); //int,float,currency
        } elseif ($num_format_type == 'n_string') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="inlineStr"><is><t>' . self::xmlspecialchars($value) . '</t></is></c>');
        } elseif ($num_format_type == 'n_auto' || 1) { //auto-detect unknown column types
            if (!is_string($value) || $value == '0' || ($value[0] != '0' && ctype_digit($value)) || preg_match("/^\-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)) {
                $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . self::xmlspecialchars($value) . '</v></c>'); //int,float,currency
            } else { //implied: ($cell_format=='string')
                $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="inlineStr"><is><t>' . self::xmlspecialchars($value) . '</t></is></c>');
            }
        }
    }

    protected function styleFontIndexes()
    {
        static $border_allowed = array('left', 'right', 'top', 'bottom');
        static $border_style_allowed = array('thin', 'medium', 'thick', 'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot');
        static $horizontal_allowed = array('general', 'left', 'right', 'justify', 'center');
        static $vertical_allowed = array('bottom', 'center', 'distributed', 'top');
        $default_font = array('size' => '10', 'name' => 'Arial', 'family' => '2');
        $fills = array('', ''); //2 placeholders for static xml later
        $fonts = array('', '', '', ''); //4 placeholders for static xml later
        $borders = array(''); //1 placeholder for static xml later
        $style_indexes = array();
        foreach ($this->cell_styles as $i => $cell_style_string) {
            $semi_colon_pos = strpos($cell_style_string, ";");
            $number_format_idx = substr($cell_style_string, 0, $semi_colon_pos);
            $style_json_string = substr($cell_style_string, $semi_colon_pos + 1);
            $style = @json_decode($style_json_string, $as_assoc = true);

            $style_indexes[$i] = array('num_fmt_idx' => $number_format_idx); //initialize entry
            if (isset($style['border']) && is_string($style['border'])) //border is a comma delimited str
            {
                $border_value['side'] = array_intersect(explode(",", $style['border']), $border_allowed);
                if (isset($style['border-style']) && in_array($style['border-style'], $border_style_allowed)) {
                    $border_value['style'] = $style['border-style'];
                }
                if (isset($style['border-color']) && is_string($style['border-color']) && $style['border-color'][0] == '#') {
                    $v = substr($style['border-color'], 1, 6);
                    $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v; // expand cf0 => ccff00
                    $border_value['color'] = "FF" . strtoupper($v);
                }
                $style_indexes[$i]['border_idx'] = self::add_to_list_get_index($borders, json_encode($border_value));
            }
            if (isset($style['fill']) && is_string($style['fill']) && $style['fill'][0] == '#') {
                $v = substr($style['fill'], 1, 6);
                $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v; // expand cf0 => ccff00
                $style_indexes[$i]['fill_idx'] = self::add_to_list_get_index($fills, "FF" . strtoupper($v));
            }
            if (isset($style['halign']) && in_array($style['halign'], $horizontal_allowed)) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['halign'] = $style['halign'];
            }
            if (isset($style['valign']) && in_array($style['valign'], $vertical_allowed)) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['valign'] = $style['valign'];
            }
            if (isset($style['wrap_text'])) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['wrap_text'] = (bool)$style['wrap_text'];
            }

            $font = $default_font;
            if (isset($style['font-size'])) {
                $font['size'] = floatval($style['font-size']); //floatval to allow "10.5" etc
            }
            if (isset($style['font']) && is_string($style['font'])) {
                if ($style['font'] == 'Comic Sans MS') {
                    $font['family'] = 4;
                }
                if ($style['font'] == 'Times New Roman') {
                    $font['family'] = 1;
                }
                if ($style['font'] == 'Courier New') {
                    $font['family'] = 3;
                }
                $font['name'] = strval($style['font']);
            }
            if (isset($style['font-style']) && is_string($style['font-style'])) {
                if (strpos($style['font-style'], 'bold') !== false) {
                    $font['bold'] = true;
                }
                if (strpos($style['font-style'], 'italic') !== false) {
                    $font['italic'] = true;
                }
                if (strpos($style['font-style'], 'strike') !== false) {
                    $font['strike'] = true;
                }
                if (strpos($style['font-style'], 'underline') !== false) {
                    $font['underline'] = true;
                }
            }
            if (isset($style['color']) && is_string($style['color']) && $style['color'][0] == '#') {
                $v = substr($style['color'], 1, 6);
                $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v; // expand cf0 => ccff00
                $font['color'] = "FF" . strtoupper($v);
            }
            if ($font != $default_font) {
                $style_indexes[$i]['font_idx'] = self::add_to_list_get_index($fonts, json_encode($font));
            }
        }
        return array('fills' => $fills, 'fonts' => $fonts, 'borders' => $borders, 'styles' => $style_indexes);
    }

    protected function writeStylesXML()
    {
        $styleIndexes = self::styleFontIndexes();
        $fills = $styleIndexes['fills'];
        $fonts = $styleIndexes['fonts'];
        $borders = $styleIndexes['borders'];
        $styleIndexes = $styleIndexes['styles'];

        // Apply custom styles set through methods
        $font = $this->font;
        $enableBorders = $this->enableBorders;
        $columnWidths = $this->columnWidths;

        $temporaryFilename = $this->tempFilename();
        $file = new XLSXWriter_BuffererWriter($temporaryFilename);

        // XML declaration and root element
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');

        // Add custom number formats
        $this->writeNumberFormats($file);

        // Add custom fonts
        $this->writeFonts($file, $fonts);

        // Add custom fills
        $this->writeFills($file, $fills);

        // Enable or disable borders based on the setting
        if ($enableBorders) {
            // Add custom borders
            $this->writeBorders($file, $borders);
        }

        // Set custom column widths
        $this->writeColumnWidths($file, $columnWidths);

        // Write cell style XF elements
        $this->writeCellStyleXF($file);

        // Write cell XF elements
        $this->writeCellXF($file, $styleIndexes);

        // Write cell styles
        $this->writeCellStyles($file);

        // Close the root element
        $file->write('</styleSheet>');

        // Close the file
        $file->close();

        return $temporaryFilename;
    }

    protected function buildAppXML()
    {
        $app_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $app_xml .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';

        $properties = array(
            'TotalTime' => '0',
            'Company' => $this->company
        );

        foreach ($properties as $tag => $value) {
            $app_xml .= '<' . $tag . '>' . self::xmlspecialchars($value) . '</' . $tag . '>';
        }

        $app_xml .= '</Properties>';
        return $app_xml;
    }

    protected function buildCoreXML()
    {
        $core_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $core_xml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';

        $core_xml .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . self::getW3CDTFDateTime() . '</dcterms:created>';

        $properties = array(
            'title' => $this->title,
            'subject' => $this->subject,
            'author' => $this->author,
            'keywords' => implode(", ", (array)$this->keywords),
            'description' => $this->description,
            'revision' => '0'
        );

        foreach ($properties as $tag => $value) {
            $core_xml .= '<dc:' . $tag . '>' . self::xmlspecialchars($value) . '</dc:' . $tag . '>';
        }

        $core_xml .= '</cp:coreProperties>';
        return $core_xml;
    }

    protected function buildRelationshipsXML()
    {
        $rels_xml = "";
        $rels_xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $rels_xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $rels_xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
        $rels_xml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
        $rels_xml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
        $rels_xml .= "\n";
        $rels_xml .= '</Relationships>';
        return $rels_xml;
    }

    protected function buildWorkbookXML()
    {
        $i = 0;
        $workbook_xml = "";
        $workbook_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $workbook_xml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $workbook_xml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
        $workbook_xml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $workbook_xml .= '<sheets>';
        foreach ($this->sheets as $sheet_name => $sheet) {
            $sheetname = self::sanitize_sheetname($sheet->sheetname);
            $workbook_xml .= '<sheet name="' . self::xmlspecialchars($sheetname) . '" sheetId="' . ($i + 1) . '" state="visible" r:id="rId' . ($i + 2) . '"/>';
            $i++;
        }
        $workbook_xml .= '</sheets>';

        // Añadir información sobre el ancho de columnas según la configuración
        $workbook_xml .= '<cols>';
        foreach ($this->columnWidths as $col => $width) {
            $workbook_xml .= '<col min="' . $col . '" max="' . $col . '" width="' . $width . '" customWidth="1"/>';
        }
        $workbook_xml .= '</cols>';

        $workbook_xml .= '<definedNames>';
        foreach ($this->sheets as $sheet_name => $sheet) {
            if ($sheet->auto_filter) {
                $sheetname = self::sanitize_sheetname($sheet->sheetname);
                $workbook_xml .= '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'' . self::xmlspecialchars($sheetname) . '\'!$A$1:' . self::xlsCell($sheet->row_count - 1, count($sheet->columns) - 1, true) . '</definedName>';
                $i++;
            }
        }
        $workbook_xml .= '</definedNames>';
        $workbook_xml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
        return $workbook_xml;
    }

    protected function buildWorkbookRelsXML()
    {
        $relationships = [];
        $relationships[] = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';

        $i = 2;
        foreach ($this->sheets as $sheet_name => $sheet) {
            $relationships[] = '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' . $sheet->xmlname . '"/>';
            $i++;
        }

        $wkbkrels_xml = '<?xml version="1.0" encoding="UTF-8"?>' . PHP_EOL;
        $wkbkrels_xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' . PHP_EOL;
        $wkbkrels_xml .= implode(PHP_EOL, $relationships) . PHP_EOL;
        $wkbkrels_xml .= '</Relationships>';

        return $wkbkrels_xml;
    }

    protected function buildContentTypesXML()
    {
        $content_types_xml = '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $content_types_xml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';

        $overrides = [
            '/_rels/.rels' => 'application/vnd.openxmlformats-package.relationships+xml',
            '/xl/_rels/workbook.xml.rels' => 'application/vnd.openxmlformats-package.relationships+xml',
            '/xl/workbook.xml' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            '/xl/styles.xml' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
            '/docProps/app.xml' => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
            '/docProps/core.xml' => 'application/vnd.openxmlformats-package.core-properties+xml',
        ];

        foreach ($this->sheets as $sheet_name => $sheet) {
            $overrides["/xl/worksheets/{$sheet->xmlname}"] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
        }

        foreach ($overrides as $partName => $contentType) {
            $content_types_xml .= '<Override PartName="' . htmlspecialchars($partName) . '" ContentType="' . htmlspecialchars($contentType) . '"/>';
        }

        $content_types_xml .= "\n";
        $content_types_xml .= '</Types>';
        return $content_types_xml;
    }

    //------------------------------------------------------------------
    /*
	 * @param $row_number int, zero based
	 * @param $column_number int, zero based
	 * @param $absolute bool
	 * @return Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
	 * */
    public static function xlsCell($row_number, $column_number, $absolute = false)
    {
        $n = $column_number;
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41) . $r;
        }
        if ($absolute) {
            return '$' . $r . '$' . ($row_number + 1);
        }
        return $r . ($row_number + 1);
    }
    //------------------------------------------------------------------
    public static function log($string)
    {
        //file_put_contents("php://stderr", date("Y-m-d H:i:s:").rtrim(is_array($string) ? json_encode($string) : $string)."\n");
        error_log(date("Y-m-d H:i:s:") . rtrim(is_array($string) ? json_encode($string) : $string) . "\n");
    }
    //------------------------------------------------------------------
    public static function sanitize_filename($filename) //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
    {
        $nonprinting = array_map('chr', range(0, 31));
        $invalid_chars = array('<', '>', '?', '"', ':', '|', '\\', '/', '*', '&');
        $all_invalids = array_merge($nonprinting, $invalid_chars);
        return str_replace($all_invalids, "", $filename);
    }
    //------------------------------------------------------------------
    public static function sanitize_sheetname($sheetname)
    {
        static $badchars  = '\\/?*:[]';
        static $goodchars = '        ';
        $sheetname = strtr($sheetname, $badchars, $goodchars);
        $sheetname = function_exists('mb_substr') ? mb_substr($sheetname, 0, 31) : substr($sheetname, 0, 31);
        $sheetname = trim(trim(trim($sheetname), "'")); //trim before and after trimming single quotes
        return !empty($sheetname) ? $sheetname : 'Sheet' . ((rand() % 900) + 100);
    }
    //------------------------------------------------------------------
    public static function xmlspecialchars($val)
    {
        //note, badchars does not include \t\n\r (\x09\x0a\x0d)
        static $badchars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
        static $goodchars = "                              ";
        return strtr(htmlspecialchars($val, ENT_QUOTES | ENT_XML1), $badchars, $goodchars); //strtr appears to be faster than str_replace
    }
    //------------------------------------------------------------------
    public static function array_first_key(array $arr)
    {
        reset($arr);
        $first_key = key($arr);
        return $first_key;
    }
    //------------------------------------------------------------------
    private static function determineNumberFormatType($num_format)
    {
        $num_format = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i", "", $num_format);
        if ($num_format == 'GENERAL') return 'n_auto';
        if ($num_format == '@') return 'n_string';
        if ($num_format == '0') return 'n_numeric';
        if (preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $num_format)) return 'n_datetime';
        if (preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $num_format)) return 'n_datetime';
        if (preg_match('/[Y]{2,4}(?![^"]*+")/i', $num_format)) return 'n_date';
        if (preg_match('/[D]{1,2}(?![^"]*+")/i', $num_format)) return 'n_date';
        if (preg_match('/[M]{1,2}(?![^"]*+")/i', $num_format)) return 'n_date';
        if (preg_match('/$(?![^"]*+")/', $num_format)) return 'n_numeric';
        if (preg_match('/%(?![^"]*+")/', $num_format)) return 'n_numeric';
        if (preg_match('/0(?![^"]*+")/', $num_format)) return 'n_numeric';
        return 'n_auto';
    }
    //------------------------------------------------------------------
    private static function numberFormatStandardized($num_format)
    {
        if ($num_format == 'money') {
            $num_format = 'dollar';
        }
        if ($num_format == 'number') {
            $num_format = 'integer';
        }

        if ($num_format == 'string')   $num_format = '@';
        else if ($num_format == 'integer')  $num_format = '0';
        else if ($num_format == 'date')     $num_format = 'YYYY-MM-DD';
        else if ($num_format == 'datetime') $num_format = 'YYYY-MM-DD HH:MM:SS';
        else if ($num_format == 'time')     $num_format = 'HH:MM:SS';
        else if ($num_format == 'price')    $num_format = '#,##0.00';
        else if ($num_format == 'dollar')   $num_format = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
        else if ($num_format == 'euro')     $num_format = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
        $ignore_until = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($num_format); $i < $ix; $i++) {
            $c = $num_format[$i];
            if ($ignore_until == '' && $c == '[')
                $ignore_until = ']';
            else if ($ignore_until == '' && $c == '"')
                $ignore_until = '"';
            else if ($ignore_until == $c)
                $ignore_until = '';
            if ($ignore_until == '' && ($c == ' ' || $c == '-'  || $c == '('  || $c == ')') && ($i == 0 || $num_format[$i - 1] != '_'))
                $escaped .= "\\" . $c;
            else
                $escaped .= $c;
        }
        return $escaped;
    }
    //------------------------------------------------------------------
    public static function add_to_list_get_index(&$haystack, $needle)
    {
        $existing_idx = array_search($needle, $haystack, $strict = true);
        if ($existing_idx === false) {
            $existing_idx = count($haystack);
            $haystack[] = $needle;
        }
        return $existing_idx;
    }
    //------------------------------------------------------------------
    public static function convert_date_time($date_input) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
    {
        $days    = 0;    # Number of days since epoch
        $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;
        $hour = $min  = $sec = 0;

        $date_time = $date_input;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches)) {
            list($junk, $year, $month, $day) = $matches;
        }
        if (preg_match("/(\d+):(\d{2}):(\d{2})/", $date_time, $matches)) {
            list($junk, $hour, $min, $sec) = $matches;
            $seconds = ($hour * 60 * 60 + $min * 60 + $sec) / (24 * 60 * 60);
        }

        //using 1900 as epoch, not 1904, ignoring 1904 special case

        # Special cases for Excel.
        if ("$year-$month-$day" == '1899-12-31')  return $seconds;    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-01-00')  return $seconds;    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-02-29')  return 60 + $seconds;    # Excel false leapday

        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch  = 1900;
        $offset = 0;
        $norm   = 300;
        $range  = $year - $epoch;

        # Set month days and check for leap year.
        $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100))) ? 1 : 0;
        $mdays = array(31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

        # Some boundary checks
        if ($year != 0 || $month != 0 || $day != 0) {
            if ($year < $epoch || $year > 9999) return 0;
            if ($month < 1     || $month > 12)  return 0;
            if ($day < 1       || $day > $mdays[$month - 1]) return 0;
        }

        # Accumulate the number of days since the epoch.
        $days = $day;    # Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1));    # Add days for past months
        $days += $range * 365;                      # Add days for past years
        $days += intval(($range) / 4);             # Add leapdays
        $days -= intval(($range + $offset) / 100); # Subtract 100 year leapdays
        $days += intval(($range + $offset + $norm) / 400);  # Add 400 year leapdays
        $days -= $leap;                                      # Already counted above

        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }

    protected function addZipDir($zip, $dir, $fileNames, $fileContents)
    {
        $zip->addEmptyDir($dir);

        foreach ($fileNames as $index => $fileName) {
            $zip->addFromString($dir . $fileName, $fileContents[$index]);
        }
    }

    protected function getSheetFileNames()
    {
        return array_map(function ($sheet) {
            return $sheet->xmlname;
        }, $this->sheets);
    }

    protected function getSheetFiles()
    {
        return array_map(function ($sheet) {
            return file_get_contents($sheet->filename);
        }, $this->sheets);
    }

    protected function getW3CDTFDateTime()
    {
        return date("Y-m-d\TH:i:s.00\Z");
    }

    protected function writeNumberFormats($file)
    {
        $file->write('<numFmts count="' . count($this->number_formats) . '">');
        foreach ($this->number_formats as $i => $v) {
            $file->write('<numFmt numFmtId="' . (164 + $i) . '" formatCode="' . self::xmlspecialchars($v) . '" />');
        }
        $file->write('</numFmts>');
    }

    protected function writeFonts($file, $fonts)
    {
        $file->write('<fonts count="' . (count($fonts) + 1) . '">');
        $file->write('<font><name val="' . htmlspecialchars($this->font['name']) . '"/><charset val="1"/><family val="2"/><sz val="' . intval($this->font['size']) . '"/></font>');
        foreach ($fonts as $font) {
            if (!empty($font)) {
                $f = json_decode($font, true);
                $file->write('<font>');
                $file->write('<name val="' . htmlspecialchars($f['name']) . '"/><charset val="1"/><family val="' . intval($f['family']) . '"/>');
                $file->write('<sz val="' . intval($f['size']) . '"/>');
                if (!empty($f['color'])) {
                    $file->write('<color rgb="' . strval($f['color']) . '"/>');
                }
                if (!empty($f['bold'])) {
                    $file->write('<b val="true"/>');
                }
                if (!empty($f['italic'])) {
                    $file->write('<i val="true"/>');
                }
                if (!empty($f['underline'])) {
                    $file->write('<u val="single"/>');
                }
                if (!empty($f['strike'])) {
                    $file->write('<strike val="true"/>');
                }
                $file->write('</font>');
            }
        }
        $file->write('</fonts>');
    }

    protected function writeFills($file, $fills)
    {
        $file->write('<fills count="' . (count($fills)) . '">');
        $file->write('<fill><patternFill patternType="none"/></fill>');
        $file->write('<fill><patternFill patternType="gray125"/></fill>');
        foreach ($fills as $fill) {
            if (!empty($fill)) {
                $file->write('<fill><patternFill patternType="solid"><fgColor rgb="' . strval($fill) . '"/><bgColor indexed="64"/></patternFill></fill>');
            }
        }
        $file->write('</fills>');
    }

    protected function writeBorders($file, $borders)
    {
        foreach ($borders as $border) {
            if (!empty($border)) {
                $pieces = json_decode($border, true);
                $border_style = !empty($pieces['style']) ? $pieces['style'] : 'hair';
                $border_color = !empty($pieces['color']) ? '<color rgb="' . strval($pieces['color']) . '"/>' : '';
                $file->write('<border diagonalDown="false" diagonalUp="false">');
                foreach (array('left', 'right', 'top', 'bottom') as $side) {
                    $show_side = in_array($side, $pieces['side']) ? true : false;
                    $file->write($show_side ? "<$side style=\"$border_style\">$border_color</$side>" : "<$side/>");
                }
                $file->write('<diagonal/>');
                $file->write('</border>');
            }
        }
    }

    protected function writeColumnWidths($file, $columnWidths)
    {
        foreach ($columnWidths as $col => $width) {
            $file->write('<col min="' . $this->columnIndexFromString($col) . '" max="' . $this->columnIndexFromString($col) . '" width="' . $width . '" customWidth="1"/>');
        }
    }

    protected function writeCellStyleXF($file)
    {
        $file->write('<cellStyleXfs count="20">');
        $file->write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
        $file->write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
        $file->write('<protection hidden="false" locked="true"/>');
        $file->write('</xf>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
        $file->write('</cellStyleXfs>');
    }

    protected function writeCellXF($file, $styleIndexes)
    {
        $file->write('<cellXfs count="' . (count($styleIndexes)) . '">');
        foreach ($styleIndexes as $v) {
            $applyAlignment = isset($v['alignment']) ? 'true' : 'false';
            $wrapText = !empty($v['wrap_text']) ? 'true' : 'false';
            $horizAlignment = isset($v['halign']) ? $v['halign'] : 'general';
            $vertAlignment = isset($v['valign']) ? $v['valign'] : 'bottom';
            $applyBorder = isset($v['border_idx']) ? 'true' : 'false';
            $applyFont = 'true';
            $borderIdx = isset($v['border_idx']) ? intval($v['border_idx']) : 0;
            $fillIdx = isset($v['fill_idx']) ? intval($v['fill_idx']) : 0;
            $fontIdx = isset($v['font_idx']) ? intval($v['font_idx']) : 0;
            $file->write('<xf applyAlignment="' . $applyAlignment . '" applyBorder="' . $applyBorder . '" applyFont="' . $applyFont . '" applyProtection="false" borderId="' . ($borderIdx) . '" fillId="' . ($fillIdx) . '" fontId="' . ($fontIdx) . '" numFmtId="' . (164 + $v['num_fmt_idx']) . '" xfId="0">');
            $file->write('	<alignment horizontal="' . $horizAlignment . '" vertical="' . $vertAlignment . '" textRotation="0" wrapText="' . $wrapText . '" indent="0" shrinkToFit="false"/>');
            $file->write('	<protection locked="true" hidden="false"/>');
            $file->write('</xf>');
        }
        $file->write('</cellXfs>');
    }

    protected function writeCellStyles($file)
    {
        $file->write('<cellStyles count="6">');
        $file->write('<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
        $file->write('<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
        $file->write('<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
        $file->write('<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
        $file->write('<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
        $file->write('<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
        $file->write('</cellStyles>');
    }

    protected function columnIndexFromString($column)
    {
        $columnIndex = 0;
        $column = strtoupper($column);

        for ($i = 0; $i < strlen($column); $i++) {
            $columnIndex = $columnIndex * 26 + ord($column[$i]) - ord('A') + 1;
        }

        return $columnIndex;
    }

    protected function writePaneAndSelection($sheet, $freezeRows = null, $freezeColumns = null)
    {
        $paneAttributes = '';
        $selectionAttributes = '';

        if ($freezeRows !== null && $freezeColumns !== null) {
            $paneAttributes = ' ySplit="' . $freezeRows . '" xSplit="' . $freezeColumns . '" topLeftCell="' . self::xlsCell($freezeRows, $freezeColumns) . '" activePane="bottomRight" state="frozen"';
            $selectionAttributes = ' activeCell="' . self::xlsCell($freezeRows, 0) . '" activeCellId="0" pane="topRight" sqref="' . self::xlsCell($freezeRows, 0) . '"';
            $selectionAttributes .= ' activeCell="' . self::xlsCell(0, $freezeColumns) . '" activeCellId="0" pane="bottomLeft" sqref="' . self::xlsCell(0, $freezeColumns) . '"';
            $selectionAttributes .= ' activeCell="' . self::xlsCell($freezeRows, $freezeColumns) . '" activeCellId="0" pane="bottomRight" sqref="' . self::xlsCell($freezeRows, $freezeColumns) . '"';
        } elseif ($freezeRows !== null) {
            $paneAttributes = ' ySplit="' . $freezeRows . '" topLeftCell="' . self::xlsCell($freezeRows, 0) . '" activePane="bottomLeft" state="frozen"';
            $selectionAttributes = ' activeCell="' . self::xlsCell($freezeRows, 0) . '" activeCellId="0" pane="bottomLeft" sqref="' . self::xlsCell($freezeRows, 0) . '"';
        } elseif ($freezeColumns !== null) {
            $paneAttributes = ' xSplit="' . $freezeColumns . '" topLeftCell="' . self::xlsCell(0, $freezeColumns) . '" activePane="topRight" state="frozen"';
            $selectionAttributes = ' activeCell="' . self::xlsCell(0, $freezeColumns) . '" activeCellId="0" pane="topRight" sqref="' . self::xlsCell(0, $freezeColumns) . '"';
        }

        $sheet->file_writer->write('<pane' . $paneAttributes . '/>');
        $sheet->file_writer->write('<selection' . $selectionAttributes . '/>');
    }

    protected function writeCol($sheet, $index, $width, $customWidth = true, $hidden = false)
    {
        $customWidthAttribute = $customWidth ? ' customWidth="true"' : '';
        $hiddenAttribute = $hidden ? ' hidden="true"' : '';

        $sheet->file_writer->write('<col collapsed="false" max="' . $index . '" min="' . $index . '" style="0"' . $customWidthAttribute . $hiddenAttribute . ' width="' . floatval($width) . '"/>');
    }

    //------------------------------------------------------------------
}

class XLSXWriter_BuffererWriter
{
    const BUFFER_LIMIT = 8191;

    protected $fd = null;
    protected $buffer = '';
    protected $check_utf8 = false;

    public function __construct($filename, $fd_fopen_flags = 'w', $check_utf8 = false)
    {
        $this->check_utf8 = $check_utf8;
        $this->fd = @fopen($filename, $fd_fopen_flags);
        if ($this->fd === false) {
            throw new Exception("Unable to open $filename for writing.");
        }
    }

    public function write($string)
    {
        $this->buffer .= $string;
        if (isset($this->buffer[self::BUFFER_LIMIT])) {
            $this->purge();
        }
    }

    public function directWrite($string)
    {
        if ($this->fd) {
            fwrite($this->fd, $string);
        }
    }

    protected function purge()
    {
        if ($this->fd) {
            if ($this->check_utf8 && !self::isValidUTF8($this->buffer)) {
                throw new Exception("Error, invalid UTF8 encoding detected.");
            }
            fwrite($this->fd, $this->buffer);
            $this->buffer = '';
        }
    }

    public function close()
    {
        $this->purge();
        if ($this->fd) {
            fclose($this->fd);
            $this->fd = null;
        }
    }

    public function __destruct()
    {
        $this->close();
    }

    public function ftell()
    {
        if ($this->fd) {
            $this->purge();
            return ftell($this->fd);
        }
        return -1;
    }

    public function fseek($pos)
    {
        if ($this->fd) {
            $this->purge();
            return fseek($this->fd, $pos);
        }
        return -1;
    }

    protected static function isValidUTF8($string)
    {
        if (function_exists('mb_check_encoding')) {
            return mb_check_encoding($string, 'UTF-8') ? true : false;
        }
        return preg_match("//u", $string) ? true : false;
    }
}
