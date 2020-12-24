<?php

namespace GentritAbazi\PhpSpreadsheetExcel\Services;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Csv as CsvWriter;
use Symfony\Component\HttpFoundation\StreamedResponse;

class PhpSpreadsheetExcelService
{
    private $columns;
    private $rows;
    private $config;
    private $spreadsheet;
    private $sheet;
    
    public function __construct($columns, $rows, $config)
    {
        $this->columns = $columns;
        $this->rows = $rows;
        $this->config = $config;
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    public function build()
    {
        $startFromColumn = 1;

        // Title
        if (isset($this->config['title'])) {
            $this->sheet->setCellValue('A1', $this->config['title'])->getStyle('A1')->getFont()->setBold(true);
            $this->sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $this->sheet->mergeCells('A1:'. $this->convertNumberToLetter(count($this->columns) + 1). '1');
            $startFromColumn += 2;
        }

        // Auto Numbering
        if (isset($this->config['auto_numbering']) && $this->config['auto_numbering'] == true) {
            $this->sheet->setCellValue('A'. $startFromColumn, 'No.');

            // Style Columns
            if (isset($this->config['columns_style'])) {
                $this->sheet->getStyle('A'. $startFromColumn)->applyFromArray($this->config['column_style']);
            }

            $columnInsideColumns = 'B';
        } else {
            $columnInsideColumns = 'A';
        }

        // Columns
        for ($i = 0; $i < count($this->columns); $i++) {
            // Column Value
            if (is_array($this->columns[$i])) {
                $this->sheet->setCellValue($columnInsideColumns. $startFromColumn, $this->columns[$i]['value']);
            } else {
                $this->sheet->setCellValue($columnInsideColumns. $startFromColumn, $this->columns[$i]);
            }

            // Style Columns
            if (isset($this->config['columns_style'])) {
                $this->sheet->getStyle($columnInsideColumns. $startFromColumn)->applyFromArray($this->config['column_style']);
            }

            // Autosize Columns
            if (isset($this->config['columns_autosize']) && !is_array($this->columns[$i])) {
                $this->sheet->getColumnDimension($columnInsideColumns)->setAutoSize(true);
            }

            // Column Settings
            if (is_array($this->columns[$i])) {
                if (isset($this->columns[$i]['width'])) {
                    $this->sheet->getColumnDimension($columnInsideColumns)->setWidth($this->columns[$i]['width']);
                }
            }

            $columnInsideColumns++;
        }

        // Auto Filter Columns
        if (isset($this->config['columns_auto_filter']) && $this->config['columns_auto_filter']) {
            $this->sheet->setAutoFilter('A'. $startFromColumn. ':'. $this->convertNumberToLetter(count($this->columns) + 1). $startFromColumn);
        }

        // Rows
        $startFromColumn += 1;
        $autoNumbering = 1;
        foreach ($this->rows as $row) {
            $columnInsideRows = 'A';
            if (isset($this->config['auto_numbering'])) {
                $this->sheet->setCellValue("$columnInsideRows$startFromColumn", $autoNumbering);
                $columnInsideRows = 'B';
            }

            foreach ($row as $val) {
                $this->sheet->setCellValue("$columnInsideRows$startFromColumn", $val);
                $columnInsideRows++;
            }

            $startFromColumn++;
            $autoNumbering++;
        }

        // Worksheet Settings
        if (isset($this->config['borders']) && $this->config['borders'] == true) {
            $worksheetDimension = $this->sheet->calculateWorksheetDimension();
            if (isset($this->config['title'])) {
                $worksheetDimension = str_ireplace('A1', 'A3', $worksheetDimension);
            }
            
            $this->sheet->getStyle($worksheetDimension)->getBorders()->applyFromArray([
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
                ]
            ]);
        }

        return true;
    }

    public function stream()
    {
        $this->build();

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
        $this->build();

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
        $this->build();
        
        $writer = new XlsxWriter($this->spreadsheet);
        $writer->save($path);

        return true;
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
