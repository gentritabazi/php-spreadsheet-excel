<?php

require_once 'vendor/autoload.php';

use GentritAbazi\PhpSpreadsheetExcel\Services\PhpSpreadsheetExcelService;

$columns = ['Column1', 'Column2', 'Column3'];
$rows = [
    ['Value For Column 1', 'Value For Column 2', 'Value For Column 3'],
    ['Text 1', 'Text 2', 'Text 3']
];
$config = [];

$spreadsheetExcel = new PhpSpreadsheetExcelService($columns, $rows, $config);
$spreadsheetExcel->save('file.xlsx');