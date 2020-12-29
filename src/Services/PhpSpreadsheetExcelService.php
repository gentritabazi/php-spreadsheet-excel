<?php

namespace GentritAbazi\PhpSpreadsheetExcel\Services;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Csv as CsvWriter;
use Symfony\Component\HttpFoundation\StreamedResponse;

class PhpSpreadsheetExcelService
{
    private $spreadsheet;
    private $totalSheets = 0;

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
    }

    public function build($activeSheet, $columns, $rows, $config)
    {
        $startFromColumn = 'A';
        $startFromColumnIndex = 1;

        // Title
        if (isset($config['title'])) {
            $activeSheet->setCellValue('A1', $config['title'])->getStyle('A1')->getFont()->setBold(true);
            $activeSheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $activeSheet->mergeCells('A1:'. $this->convertNumberToLetter(count($columns) + 1). '1');
            $startFromColumnIndex += 2;
        }

        // Auto Numbering
        if (isset($config['auto_numbering']) && $config['auto_numbering'] == true) {
            $activeSheet->setCellValue('A'. $startFromColumnIndex, 'No.');

            // Style Columns
            if (isset($config['columns_style'])) {
                $activeSheet->getStyle('A'. $startFromColumnIndex)->applyFromArray($config['column_style']);
            }

            $startFromColumn++;
        }

        // Columns
        for ($i = 0; $i < count($columns); $i++) {
            // Column Value
            if (is_array($columns[$i])) {
                $activeSheet->setCellValue($startFromColumn. $startFromColumnIndex, $columns[$i]['value']);
            } else {
                $activeSheet->setCellValue($startFromColumn. $startFromColumnIndex, $columns[$i]);
            }

            // Style Columns
            if (isset($config['columns_style'])) {
                $activeSheet->getStyle($startFromColumn. $startFromColumnIndex)->applyFromArray($config['column_style']);
            }

            // Autosize Columns
            if (isset($config['columns_autosize']) && !is_array($columns[$i])) {
                $activeSheet->getColumnDimension($startFromColumn)->setAutoSize(true);
            }

            // Column Settings
            if (is_array($columns[$i])) {
                if (isset($columns[$i]['width'])) {
                    $activeSheet->getColumnDimension($startFromColumn)->setWidth($columns[$i]['width']);
                }
            }

            $startFromColumn++;
        }

        // Auto Filter Columns
        if (isset($config['columns_auto_filter']) && $config['columns_auto_filter']) {
            $activeSheet->setAutoFilter('A'. $startFromColumnIndex. ':'. $this->convertNumberToLetter(count($columns) + 1). $startFromColumnIndex);
        }

        // Rows
        $startFromColumnIndex += 1;
        $autoNumbering = 1;
        foreach ($rows as $row) {
            $columnInsideRows = 'A';
            if (isset($config['auto_numbering'])) {
                $activeSheet->setCellValue("$columnInsideRows$startFromColumnIndex", $autoNumbering);
                $columnInsideRows = 'B';
            }

            foreach ($row as $val) {
                $activeSheet->setCellValue("$columnInsideRows$startFromColumnIndex", $val);
                $columnInsideRows++;
            }

            $startFromColumnIndex++;
            $autoNumbering++;
        }

        // Worksheet Settings
        if (isset($config['borders']) && $config['borders'] == true) {
            $worksheetDimension = $activeSheet->calculateWorksheetDimension();
            if (isset($config['title'])) {
                $worksheetDimension = str_ireplace('A1', 'A3', $worksheetDimension);
            }
            
            $activeSheet->getStyle($worksheetDimension)->getBorders()->applyFromArray([
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
                ]
            ]);
        }

        return true;
    }

    public function createSheet($sheetName, $columns, $rows, $config)
    {
        $this->totalSheets += 1;

        if ($this->totalSheets > 1) {
            $this->spreadsheet->createSheet();
        }
        
        $this->spreadsheet->setActiveSheetIndex($this->totalSheets - 1);

        $activeSheet = $this->spreadsheet->getActiveSheet();

        $activeSheet->setTitle($sheetName);

        $this->build($activeSheet, $columns, $rows, $config);
    }

    public function stream()
    {
        $writer = new XlsxWriter($this->spreadsheet);
        $headers = [
            'Cache-Control' => 'must-revalidate, post-check=0, pre-check=0',
            'Content-type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition' => 'attachment; filename=file.xlsx',
            'Expires' => '0',
            'Pragma' => 'public'
        ];

        $response = new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            }
        );

        foreach ($headers as $key => $value) {
            $response->headers->set($key, $value);
        }

        return $response;
    }

    public function streamAsCsv()
    {
        $writer = new CsvWriter($this->spreadsheet);

        $headers = [
            'Cache-Control' => 'must-revalidate, post-check=0, pre-check=0',
            'Content-type' => 'text/csv',
            'Content-Disposition' => 'attachment; filename=file.csv',
            'Expires' => '0',
            'Pragma' => 'public'
        ];

        $response = new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            }
        );

        foreach ($headers as $key => $value) {
            $response->headers->set($key, $value);
        }

        return $response;
    }

    public function save($path)
    {
        $writer = new XlsxWriter($this->spreadsheet);
        $writer->save($path);

        return true;
    }

    public function download($fileName = 'file.xlsx')
    {
        $writer = new XlsxWriter($this->spreadsheet);

        $headers = [
            'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition' => 'attachment; filename="'. $fileName.'"'
        ];
        
        foreach ($headers as $key => $value) {
            header($key. ': '. $value);
        }

        $writer->save('php://output');
    }

    public function downloadAsCsv($fileName = 'file.csv')
    {
        $writer = new XlsxWriter($this->spreadsheet);

        $headers = [
            'Content-Type' => 'text/csv',
            'Content-Disposition' => 'attachment; filename="'. $fileName.'"'
        ];
        
        foreach ($headers as $key => $value) {
            header($key. ': '. $value);
        }

        $writer->save('php://output');
    }

    private function convertNumberToLetter($number)
    {
        $startFrom = 'A';

        for ($i = 0; $i < $number; $i++) {
            if ($i == ($number - 1)) {
                return $startFrom;
            }

            $startFrom++;
        }
    }
}
