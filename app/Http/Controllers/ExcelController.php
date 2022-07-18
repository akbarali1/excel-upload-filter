<?php

namespace App\Http\Controllers;

use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;

/**
 * Created by PhpStorm.
 * Filename: ExcelController.php
 * Project Name: excelfiler.loc
 * Author: Акбарали
 * Date: 15/07/2022
 * Time: 20:40
 * Github: https://github.com/akbarali1
 * Telegram: @akbar_aka
 * E-mail: me@akbarali.uz
 */
class ExcelController extends Controller
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

    public function generateExcelDownload($file = 'helloWorld.xlsx')
    {
        $spreadsheet = new Spreadsheet();
        $sheet       = $spreadsheet->getActiveSheet();

        $this->excelHeaderGenerate($sheet);
        $this->excelBodyGenerate($spreadsheet);

        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;
    }

    public function excelHeaderGenerate($sheet)
    {
        $ont_table = [
            [
                'name'  => '№',
                'table' => "A",
            ],
            [
                'name'  => 'Название',
                'table' => "B",
            ],
            [
                'name'  => 'Код товара',
                'table' => "C",
            ],
            [
                'name'  => 'Цена',
                'table' => "D",
            ],
            [
                'name'  => 'Артикул',
                'table' => "E",
            ],
            [
                'name'  => 'Количество',
                'table' => "F",
            ],
            [
                'name'  => 'Цена в USD',
                'table' => "G",
            ],
            [
                'name'  => 'Цена по умолчанию',
                'table' => "H",
            ],
            [
                'name'  => 'Поставщик',
                'table' => "I",
            ],
            [
                'name'  => 'Цена поставщика',
                'table' => "J",
            ],
        ];

        foreach ($ont_table as $key => $value) {
            $sheet->setCellValue($value['table'].'1', $value['name']);
            $sheet->getStyle($value['table'].'1')->getBorders()
                ->getOutline()
                ->setBorderStyle(Border::BORDER_THICK)
                ->setColor(new Color('000000'));
            $sheet->getStyle($value['table'].'1')
                ->getAlignment()
                ->setWrapText(true);
        }

        return $sheet;

    }

    public function excelBodyGenerate($spreadsheet)
    {
        $this->createSheetData($spreadsheet, 'database');


        $spreadsheet->setActiveSheetIndex(0);

        return $spreadsheet;

    }

    private function createSheetData(Spreadsheet $spreadsheet, string $rule): void
    {
        $styleArray = [
            'borders'   => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'alignment' => [
                'wrapText' => true,
            ],
            'font'      => [
                //'size' => 14,
            ],
        ];
        $sheet      = $spreadsheet->createSheet()->setTitle($rule);
        $this->productCreate($sheet, $styleArray);
        $this->supplierCreate($sheet, $styleArray);

        //Data base sahifasini himoyalash
        $sheet->getProtection()->setPassword('akbarali');
        $sheet->getProtection()->setSheet(true); // This should be enabled in order to enable any of the following!
        $sheet->getProtection()->setSort(true);
        $sheet->getProtection()->setInsertRows(true);
        $sheet->getProtection()->setFormatCells(true);
    }

    private function productCreate($sheet, $styleArray): void
    {
        $products = DB::table('products')->whereNull('deleted_at')->orderBy('sku')->select('id', 'name')->get()->transform(function ($item) {
            $name = json_decode($item->name, true);

            return [
                'id'   => $item->id,
                'name' => $name['uz'],
            ];
        })->toArray();
        $products = collect($products)->sortBy('name')->toArray();
        $sheet->setCellValue('A1', 'product_name')->getStyle('A1')->applyFromArray($styleArray);
        $sheet->setCellValue('B1', 'product_id')->getStyle('B1')->applyFromArray($styleArray);

        $i = 1;
        foreach ($products as $value) {
            $i++;
            $sheet->getStyle('A'.$i.':B'.$i)->applyFromArray($styleArray);
            $sheet->setCellValue('A'.$i, $value['name'])->getColumnDimension('A')->setWidth(85);
            $sheet->setCellValue('B'.$i, $value['id'])->getColumnDimension('B')->setWidth(11);
        }

    }

    private function supplierCreate($sheet, $styleArray): void
    {
        $suppliers = DB::table('suppliers')->whereNull('deleted_at')->orderBy('id')->select('id', 'name')->get()->transform(function ($item) {
            return [
                'id'   => $item->id,
                'name' => $item->name,
            ];
        })->toArray();
        $sheet->setCellValue('D1', 'supplier_name')->getStyle('D1')->applyFromArray($styleArray);
        $sheet->setCellValue('E1', 'supplier_id')->getStyle('E1')->applyFromArray($styleArray);

        $i = 1;
        foreach ($suppliers as $value) {
            $i++;
            $sheet->getStyle('D'.$i.':E'.$i)->applyFromArray($styleArray);
            $sheet->setCellValue('D'.$i, $value['name'])->getColumnDimension('D')->setWidth(30);
            $sheet->setCellValue('E'.$i, $value['id'])->getColumnDimension('E')->setWidth(12);
        }

    }


}
