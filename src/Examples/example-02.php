<?php

require_once 'vendor/autoload.php';

use GentritAbazi\PhpSpreadsheetExcel\Services\PhpSpreadsheetExcelService;

$columns = ['Column1', 'Column2', 'Column3'];
$rows = [
    ['Value For Column 1', 'Value For Column 2', 'Value For Column 3'],
    ['Text 1', 'Text 2', 'Text 3']
];
$config = [
    'title' => 'Hello World',
    'auto_numbering' => true,
    'columns_autosize' => true,
    'columns_auto_filter' => true,
    'borders' => true
];

$spreadsheetExcel = new PhpSpreadsheetExcelService();
$spreadsheetExcel->createSheet('Sheet 1', $columns, $rows, $config);
$spreadsheetExcel->save('file.xlsx');