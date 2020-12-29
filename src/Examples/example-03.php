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
        $rows = User::get()->toArray();
        $config = ['columns_autosize' => true];

        $spreadsheetExcel = new PhpSpreadsheetExcelService();
        $spreadsheetExcel->createSheet('Sheet 1', $columns, $rows, $config);
        $spreadsheetExcel->download('file.xlsx');
    }
}
