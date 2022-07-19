<?php

namespace App\Http\Controllers;

use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Protection;

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
        $styleArray  = [
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


        $this->excelHeaderGenerate($sheet, $styleArray);
        $this->excelBodyGenerate($spreadsheet, $styleArray);

        $writer = new Xlsx($spreadsheet);
        $writer->setPreCalculateFormulas(false);
        $writer->save($file);

        return $file;
    }

    public function excelHeaderGenerate($sheet, $styleArray)
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
            //            [
            //                'name'  => 'Артикул',
            //                'table' => "E",
            //            ],
            [
                'name'  => 'Количество',
                'table' => "E",
            ],
            [
                'name'  => 'Цена в USD',
                'table' => "F",
            ],
            [
                'name'  => 'Цена по умолчанию',
                'table' => "G",
            ],
            [
                'name'  => 'Поставщик',
                'table' => "H",
            ],
            [
                'name'  => 'Код Поставщик',
                'table' => "I",
            ],
        ];

        foreach ($ont_table as $value) {
            $sheet->setCellValue($value['table'].'1', $value['name'])
                ->getStyle($value['table'].'1')->applyFromArray($styleArray);
        }

        return $sheet;

    }

    public function excelBodyGenerate($spreadsheet, $styleArray)
    {
        $this->createSheetData($spreadsheet, 'product', $styleArray);
        $this->createSheetData($spreadsheet, 'supplier', $styleArray);
        $spreadsheet->setActiveSheetIndex(0);
        $sheet = $spreadsheet->getActiveSheet();
        //        $sheetDatabase = $spreadsheet->getActiveSheet(1);
        $this->workSheeetGenerate($sheet, $styleArray);

        //Hamma tablitsalarni himoyalaymiz
        $sheet->getProtection()->setSheet(true);
        //Taxrirlash kerak bo'lganlarni ochamiz
        $sheet->getStyle('B2:I101')->getProtection()->setLocked(Protection::PROTECTION_UNPROTECTED);

        return $spreadsheet;

    }

    private function workSheeetGenerate($sheet, $styleArray)
    {
        $limit         = 102;
        $g             = 0;
        $product_count = DB::table('products')->whereNull('deleted_at')->count('id') + 1;
        for ($i = 2; $i < $limit; $i++) {
            $g++;
            $sheet->getStyle('A'.$i.':I'.$i)->applyFromArray($styleArray);
            $sheet->setCellValue('A'.$i, $g.')')->getColumnDimension('A')->setWidth(6);

            $sheet->getColumnDimension('B')->setWidth(40);
            $validation = $sheet->getCell('B'.$i)->getDataValidation();
            $validation->setType(DataValidation::TYPE_LIST);
            $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
            $validation->setAllowBlank(false);
            $validation->setShowInputMessage(true);
            $validation->setShowErrorMessage(true);
            $validation->setShowDropDown(true);
            $validation->setErrorTitle('Xatolik');
            $validation->setError('Bu qiymat bazada yo`q.');
            $validation->setPromptTitle('Productni tanlang');
            $validation->setPrompt('Iltimos birorta productni birini tanlang.');
            $validation->setFormula1('=product!B2:B'.$product_count);
            // C tablitsa
            $formula = '=IFERROR(VLOOKUP(B'.$i.',product!B2:C'.$product_count.',2,FALSE),0)';
            $sheet->setCellValue('C'.$i, $formula);

        }
    }

    private function createSheetData(Spreadsheet $spreadsheet, string $rule, $styleArray): void
    {
        $sheet = $spreadsheet->createSheet()->setTitle($rule);
        if ($rule === 'product') {
            $this->productCreate($sheet, $styleArray);
        } elseif ($rule === 'supplier') {
            $this->supplierCreate($sheet, $styleArray);
        }

        //Sahifani himoyalash
        $sheet->getProtection()->setPassword('akbarali');
        $sheet->getProtection()->setSheet(true); // This should be enabled in order to enable any of the following!
        $sheet->getProtection()->setSort(true);
        $sheet->getProtection()->setInsertRows(true);
        $sheet->getProtection()->setFormatCells(true);

    }

    private function productCreate($sheet, $styleArray): void
    {
        $products = DB::table('products')->whereNull('deleted_at')->orderBy('sku')->select('id', 'name', 'supplier_id')->get()->transform(function ($item) {
            $name = json_decode($item->name, true);

            return [
                'id'          => $item->id,
                'name'        => $name['uz'],
                'supplier_id' => $item->supplier_id,
            ];
        })->toArray();
        $products = collect($products)->sortBy('name')->toArray();
        $sheet->setCellValue('A1', '№')->getStyle('A1')->applyFromArray($styleArray);
        $sheet->setCellValue('B1', 'product_name')->getStyle('B1')->applyFromArray($styleArray);
        $sheet->setCellValue('C1', 'product_id')->getStyle('C1')->applyFromArray($styleArray);
        $sheet->setCellValue('D1', 'supplier_id')->getStyle('D1')->applyFromArray($styleArray);
        $i = 1;
        $g = 0;
        foreach ($products as $value) {
            $i++;
            $g++;
            $sheet->getStyle('A'.$i.':D'.$i)->applyFromArray($styleArray);
            $sheet->setCellValue('A'.$i, $g.')')->getColumnDimension('A')->setWidth(6);
            $sheet->setCellValue('B'.$i, $value['name'])->getColumnDimension('B')->setWidth(85);
            $sheet->setCellValue('C'.$i, $value['id'])->getColumnDimension('C')->setWidth(11);
            $sheet->setCellValue('D'.$i, $value['supplier_id'])->getColumnDimension('D')->setWidth(11);
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
        $sheet->setCellValue('A1', '№')->getStyle('B1')->applyFromArray($styleArray);
        $sheet->setCellValue('B1', 'supplier_name')->getStyle('B1')->applyFromArray($styleArray);
        $sheet->setCellValue('C1', 'supplier_id')->getStyle('C1')->applyFromArray($styleArray);

        $i = 1;
        $g = 0;
        foreach ($suppliers as $value) {
            $i++;
            $g++;
            $sheet->getStyle('A'.$i.':C'.$i)->applyFromArray($styleArray);
            $sheet->setCellValue('A'.$i, $g.')')->getColumnDimension('A')->setWidth(5);
            $sheet->setCellValue('B'.$i, $value['name'])->getColumnDimension('B')->setWidth(30);
            $sheet->setCellValue('C'.$i, $value['id'])->getColumnDimension('C')->setWidth(12);
        }

    }


}
