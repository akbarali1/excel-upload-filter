<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
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

    public function upload(Request $request)
    {
        $file = $request->file('file');

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($file);

        $sheet = $spreadsheet->getSheet(0)->toArray();

        $one = $sheet[0];
        unset($sheet[0]);

        $array = [];
        foreach ($sheet as $row) {
            if (!is_null($row[1])) {
                $array[] = array_combine($one, $row);
            }
        }

        return response()->json($array, 200, [], JSON_UNESCAPED_UNICODE);

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
                'with'  => 5,
            ],
            [
                'name'  => 'Название',
                'table' => "B",
                'with'  => 35,
            ],
            [
                'name'  => 'Код товара',
                'table' => "C",
                'with'  => 12,
            ],
            [
                'name'  => 'Цена',
                'table' => "D",
                'with'  => 10,
            ],
            //            [
            //                'name'  => 'Артикул',
            //                'table' => "E",
            //            ],
            [
                'name'  => 'Количество',
                'table' => "E",
                'with'  => 12,
            ],
            [
                'name'  => 'Цена в USD',
                'table' => "F",
                'with'  => 12,
            ],
            [
                'name'   => 'Поставщик',
                'table'  => "G",
                'with'   => 20,
                'filter' => true,
            ],
            [
                'name'  => 'Код Поставщик',
                'table' => "H",
                'with'  => 15,
            ],
            [
                'name'  => 'ДП 1',
                'table' => "I",
                'with'  => 20,
            ],
            [
                'name'  => 'ДП ID 1',
                'table' => "J",
                'with'  => 8,
            ],
            [
                'name'  => 'ДП Цена 1',
                'table' => "K",
                'with'  => 10,
            ],

            [
                'name'  => 'ДП 2',
                'table' => "L",
                'with'  => 20,
            ],
            [
                'name'  => 'ДП ID 2',
                'table' => "M",
                'with'  => 8,
            ],
            [
                'name'  => 'ДП Цена 2',
                'table' => "N",
                'with'  => 10,
            ],
            [
                'name'  => 'ДП 3',
                'table' => "O",
                'with'  => 20,
            ],
            [
                'name'  => 'ДП ID 3',
                'table' => "P",
                'with'  => 8,
            ],
            [
                'name'  => 'ДП Цена 3',
                'table' => "Q",
                'with'  => 10,
            ],
            [
                'name'  => 'ДП 4',
                'table' => "R",
                'with'  => 20,
            ],
            [
                'name'  => 'ДП ID 4',
                'table' => "S",
                'with'  => 8,
            ],
            [
                'name'  => 'ДП Цена 4',
                'table' => "T",
                'with'  => 10,
            ],
            [
                'name'  => 'ДП 5',
                'table' => "U",
                'with'  => 20,
            ],
            [
                'name'  => 'ДП ID 5',
                'table' => "V",
                'with'  => 8,
            ],
            [
                'name'  => 'ДП Цена 5',
                'table' => "W",
                'with'  => 10,
            ],
        ];

        foreach ($ont_table as $value) {

            $sheet->setCellValue($value['table'].'1', $value['name'])->getColumnDimension($value['table'])->setWidth($value['with']);
            $sheet->setCellValue($value['table'].'1', $value['name'])
                ->getStyle($value['table'].'1')->applyFromArray($styleArray);

            if (isset($value['filter'])) {
                $sheet->setAutoFilter($value['table'].'1:'.$value['table'].'101');
                //                $autoFilter   = $sheet->getAutoFilter();
                //                $columnFilter = $autoFilter->getColumn($value['table']);
                //                $columnFilter->setFilterType(
                //                    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
                //                );
            }
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
        foreach ($this->editAccessColumn() as $value) {
            $sheet->getStyle($value['column'].$value['row'].":".$value['column'].$value['end'])->getProtection()->setLocked(Protection::PROTECTION_UNPROTECTED);
        }

        return $spreadsheet;

    }

    private function editAccessColumn(): array
    {
        return [
            [
                'column' => 'B',
                'row'    => '2',
                'end'    => '101',
            ],
            [
                'column' => 'D',
                'row'    => '2',
                'end'    => '101',
            ],
            [
                'column' => 'E',
                'row'    => '2',
                'end'    => '101',
            ],
            [
                'column' => 'F',
                'row'    => '2',
                'end'    => '101',
            ],
            [
                'column' => 'G',
                'row'    => '2',
                'end'    => '101',
                'other'  => 'H',
            ],
            [
                'column' => 'I',
                'row'    => '2',
                'end'    => '101',
                'other'  => 'J',
            ],
            [
                'column' => 'L',
                'row'    => '2',
                'end'    => '101',
                'other'  => 'M',
            ],
            [
                'column' => 'O',
                'row'    => '2',
                'end'    => '101',
                'other'  => 'P',
            ],
            [
                'column' => 'R',
                'row'    => '2',
                'end'    => '101',
                'other'  => 'S',
            ],
            [
                'column' => 'U',
                'row'    => '2',
                'end'    => '101',
                'other'  => 'V',
            ],
        ];
    }

    private function workSheeetGenerate($sheet, $styleArray)
    {
        $limit          = 102;
        $g              = 0;
        $product_count  = DB::table('products')->whereNull('deleted_at')->count('id') + 1;
        $supplier_count = DB::table('suppliers')->whereNull('deleted_at')->count('id') + 1;
        for ($i = 2; $i < $limit; $i++) {
            $g++;
            $sheet->getStyle('A'.$i.':W'.$i)->applyFromArray($styleArray);
            $sheet->setCellValue('A'.$i, $g.')')->getColumnDimension('A')->setWidth(6);

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
            // D tablitsa Priceni tashlash
            // $sheet->setCellValue('D'.$i, '=IFERROR(VLOOKUP(B'.$i.',product!B2:E'.$product_count.',4,FALSE),0)');
            //G tablitsa Postavchikni tanlang
            foreach ($this->editAccessColumn() as $value) {
                if (isset($value['other'])) {
                    $this->supplierEditAccess($sheet, $supplier_count, $value['column'], $value['other'], $i, $styleArray);
                }
            }

        }
    }

    private function supplierEditAccess($sheet, $supplier_count, $column, $edit_access_column, $i, $styleArray): void
    {
        $validation = $sheet->getCell($column.$i)->getDataValidation();
        $validation->setType(DataValidation::TYPE_LIST);
        $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
        $validation->setAllowBlank(false);
        $validation->setShowInputMessage(true);
        $validation->setShowErrorMessage(true);
        $validation->setShowDropDown(true);
        $validation->setErrorTitle('Xatolik');
        $validation->setError('Bu qiymat bazada yo`q.');
        $validation->setPromptTitle('Postavshikni tanlang');
        $validation->setPrompt('Iltimos birorta productni birini tanlang.');
        $validation->setFormula1('=supplier!B2:B'.$supplier_count);

        //H tablitsa Supplier ID
        $formula_supplier = '=IFERROR(VLOOKUP('.$column.$i.',supplier!B2:C'.$supplier_count.',2,FALSE),0)';
        $sheet->setCellValue($edit_access_column.$i, $formula_supplier);

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
        $products = DB::table('products')->whereNull('deleted_at')->orderBy('sku')->select('id', 'name', 'supplier_id', 'price')->get()->transform(function ($item) {
            $name = json_decode($item->name, true);

            return [
                'id'          => $item->id,
                'name'        => $name['uz'],
                'supplier_id' => $item->supplier_id,
                'price'       => $item->price,
            ];
        })->toArray();
        $products = collect($products)->sortBy('name')->toArray();
        $sheet->setCellValue('A1', '№')->getStyle('A1')->applyFromArray($styleArray);
        $sheet->setCellValue('B1', 'product_name')->getStyle('B1')->applyFromArray($styleArray);
        $sheet->setCellValue('C1', 'product_id')->getStyle('C1')->applyFromArray($styleArray);
        $sheet->setCellValue('D1', 'supplier_id')->getStyle('D1')->applyFromArray($styleArray);
        $sheet->setCellValue('E1', 'price')->getStyle('E1')->applyFromArray($styleArray);
        $i = 1;
        $g = 0;
        foreach ($products as $value) {
            $i++;
            $g++;
            $sheet->getStyle('A'.$i.':E'.$i)->applyFromArray($styleArray);
            $sheet->setCellValue('A'.$i, $g.')')->getColumnDimension('A')->setWidth(6);
            $sheet->setCellValue('B'.$i, $value['name'])->getColumnDimension('B')->setWidth(85);
            $sheet->setCellValue('C'.$i, $value['id'])->getColumnDimension('C')->setWidth(11);
            $sheet->setCellValue('D'.$i, $value['supplier_id'])->getColumnDimension('D')->setWidth(11);
            $sheet->setCellValue('E'.$i, $value['price'])->getColumnDimension('E')->setWidth(10);
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
