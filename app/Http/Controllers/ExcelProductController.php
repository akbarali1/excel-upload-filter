<?php

namespace App\Http\Controllers;

use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * Created by PhpStorm.
 * Filename: ExcelProductController.php
 * Project Name: excelfiler.loc
 * Author: Акбарали
 * Date: 16/07/2022
 * Time: 11:37
 * Github: https://github.com/akbarali1
 * Telegram: @akbar_aka
 * E-mail: me@akbarali.uz
 */
class ExcelProductController extends Controller
{

    public function index()
    {
        return view('excel.index');
    }

    public function download()
    {
        $file     = $this->generateExcelDownload();
        $content  = file_get_contents($file);
        $response = response($content, 200)->header('Content-Type', 'application/vnd.ms-excel');
        $response->header('Content-Disposition', 'attachment; filename="'.$file.'"');
        unlink($file);

        return $response;

    }


    public function generateExcelDownload($file = 'product.xlsx')
    {
        $spreadsheet = new Spreadsheet();

        // Add some data
        $spreadsheet->setActiveSheetIndex(0);
        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet
            ->setCellValue('A1', 'Product')
            ->setCellValue('B1', 'Quantity')
            ->setCellValue('C1', 'Unit Price')
            ->setCellValue('D1', 'Price')
            ->setCellValue('E1', 'VAT')
            ->setCellValue('F1', 'Total');

        // Define named formula
        $spreadsheet->addNamedFormula(new \PhpOffice\PhpSpreadsheet\NamedFormula('GERMAN_VAT_RATE', $worksheet, '=16.0%'));
        $spreadsheet->addNamedFormula(new \PhpOffice\PhpSpreadsheet\NamedFormula('CALCULATED_PRICE', $worksheet, '=$B1*$C1'));
        $spreadsheet->addNamedFormula(new \PhpOffice\PhpSpreadsheet\NamedFormula('GERMAN_VAT', $worksheet, '=$D1*GERMAN_VAT_RATE'));
        $spreadsheet->addNamedFormula(new \PhpOffice\PhpSpreadsheet\NamedFormula('TOTAL_INCLUDING_VAT', $worksheet, '=$D1+$E1'));

        $worksheet
            ->setCellValue('A2', 'Advanced Web Application Architecture')
            ->setCellValue('B2', 2)
            ->setCellValue('C2', 23.0)
            ->setCellValue('D2', '=CALCULATED_PRICE')
            ->setCellValue('E2', '=GERMAN_VAT')
            ->setCellValue('F2', '=TOTAL_INCLUDING_VAT');
        $spreadsheet->getActiveSheet()
            ->setCellValue('A3', 'Object Design Style Guide')
            ->setCellValue('B3', 5)
            ->setCellValue('C3', 12.0)
            ->setCellValue('D3', '=CALCULATED_PRICE')
            ->setCellValue('E3', '=GERMAN_VAT')
            ->setCellValue('F3', '=TOTAL_INCLUDING_VAT');
        $spreadsheet->getActiveSheet()
            ->setCellValue('A4', 'PHP For the Web')
            ->setCellValue('B4', 3)
            ->setCellValue('C4', 10.0)
            ->setCellValue('D4', '=CALCULATED_PRICE')
            ->setCellValue('E4', '=GERMAN_VAT')
            ->setCellValue('F4', '=TOTAL_INCLUDING_VAT');

        // Use a relative named range to provide the totals for rows 2-4
        $spreadsheet->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('COLUMN_TOTAL', $worksheet, '=A$2:A$4'));

        $spreadsheet->getActiveSheet()
            ->setCellValue('B6', '=SUBTOTAL(109,COLUMN_TOTAL)')
            ->setCellValue('D6', '=SUBTOTAL(109,COLUMN_TOTAL)')
            ->setCellValue('E6', '=SUBTOTAL(109,COLUMN_TOTAL)')
            ->setCellValue('F6', '=SUBTOTAL(109,COLUMN_TOTAL)');


        $spreadsheet->getSheet(0);
        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;


        //        $spreadsheet = new Spreadsheet();
        //        $sheet       = $spreadsheet->getActiveSheet();
        //        $sheet->setCellValue('A1', 'Hello World !');
        //
        //        $writer = new Xlsx($spreadsheet);
        //        $writer->save($filename);
        //
        //        return $filename;
    }


}
