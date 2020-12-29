<?php

namespace App\Http\Controllers;

use App\User;
use App\Http\Controllers\Controller;
use GentritAbazi\PhpSpreadsheetExcel\Services\PhpSpreadsheetExcelService;
    
class TestController extends Controller
{
    public function index()
    {
        $columns = ['Id', 'Name', 'Email'];
        $rows = User::get();
        $config = ['columns_autosize' => true];

        $parsedRows = $rows->map(function($row) {
            $parsedRow['id'] = ['value' => $row['id']];
            $parsedRow['name'] = ['value' => $row['name']];
            $parsedRow['email'] = ['value' => $row['email']];

            // Style row - https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#styles
            if($row['name'] == 'Gentrit') {
                $parsedRow['name']['style'] = [
                    'font' => ['color' => ['rgb' => 'FF0000']]
                ];
            }
            
            return $parsedRow;
        });

        $spreadsheetExcel = new PhpSpreadsheetExcelService();
        $spreadsheetExcel->createSheet('Sheet 1', $columns, $parsedRows, $config);
        $spreadsheetExcel->download('file.xlsx');
    }
}
