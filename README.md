### Introduction

**PhpSpreadsheetExcel** is a package thats offers you to build Excel fast and easy in PHP.

This package depends on [PHPOffice/PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet).

### Installation
* composer require **gentritabazi01/phpspreadsheetexcel**.

### Usage

[Example 1](src/Examples/example-01.php):

```php
<?php

use GentritAbazi\PhpSpreadsheetExcel\Services\PhpSpreadsheetExcelService;

$columns = ['Column1', 'Column2', 'Column3'];
$rows = [
    ['Value For Column 1', 'Value For Column 2', 'Value For Column 3'],
    ['Text 1', 'Text 2', 'Text 3']
];
$config = [];

$spreadsheetExcel = new PhpSpreadsheetExcelService();
$spreadsheetExcel->createSheet('Sheet 1', $columns, $rows, $config);
$spreadsheetExcel->save('file.xlsx');
```

Output:

![...](screenshots/example-01.png)


[Example 2](src/Examples/example-02.php):

```php
<?php

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
```

Output:

![...](screenshots/example-02.png)

[Example 3](src/Examples/example-03.php):

```php
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
```

Output:

![...](screenshots/example-03.png)

[Example 4](src/Examples/example-04.php):

```php
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
```

Output:

![...](screenshots/example-04.png)

### Available Functions
* stream.
* streamAsCsv.
* save.
* createSheet.
* download($fileName | 'file.xlsx').
* downloadAsCsv($fileName | 'file.csv').
