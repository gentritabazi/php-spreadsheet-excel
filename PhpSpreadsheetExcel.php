<?php

namespace Infrastructure\Services;

use PhpOffice\PhpSpreadsheet\Writer\Csv as CsvWriter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Symfony\Component\HttpFoundation\StreamedResponse;

class PhpSpreadsheetExcel
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
            $this->sheet->mergeCells('A1:'. excelPartNrToLetter(count($this->columns) + 1). '1');
            $startFromColumn += 2;
        }

        // No.
        if (!isset($this->config['disable_no'])) {
            $this->sheet->setCellValue('A'. $startFromColumn, 'No');
            $columnInsideColumns = 'B';
        } else {
            $columnInsideColumns = 'A';
        }

        // Columns
        for ($i = 0; $i < count($this->columns); $i++) {
            if (is_array($this->columns[$i])) {
                $this->sheet->setCellValue($columnInsideColumns. $startFromColumn, $this->columns[$i]['value']);
            } else {
                $this->sheet->setCellValue($columnInsideColumns. $startFromColumn, $this->columns[$i]);
            }

            // Autosize
            if (isset($this->config['autosize']) && !is_array($this->columns[$i])) {
                $this->sheet->getColumnDimension($columnInsideColumns)->setAutoSize(true);
            }

            // Settingss
            if (is_array($this->columns[$i])) {
                if (isset($this->columns[$i]['width'])) {
                    $this->sheet->getColumnDimension($columnInsideColumns)->setWidth($this->columns[$i]['width']);
                }
            }

            $columnInsideColumns++;
        }

        // Rows
        $startFromColumn += 1;
        $rendNo = 1;
        foreach ($this->rows as $row) {
            $columnInsideRows = 'A';
            if (!isset($this->config['disable_no'])) {
                $this->sheet->setCellValue("$columnInsideRows$startFromColumn", $rendNo);
                $columnInsideRows = 'B';
            }

            foreach ($row as $val) {
                $this->sheet->setCellValue("$columnInsideRows$startFromColumn", $val);
                $columnInsideRows++;
            }

            $startFromColumn++;
            $rendNo++;
        }

        if (isset($this->config['set_borders'])) {
            $this->sheet->getStyle(str_ireplace('A1', 'A3', $this->sheet->calculateWorksheetDimension()))->getBorders()->applyFromArray([
            'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
                ]
            ]);
        }

        return true;
    }

    public function stream()
    {
        $writer = new Xlsx($this->spreadsheet);
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
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($path);

        return true;
    }
}
