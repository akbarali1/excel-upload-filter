<?php

namespace App\Http\Controllers;

use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
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

        $products = DB::table('products')->limit(100)->get()->transform(function ($item) {
            return [
                'id'    => $item->id,
                'name'  => json_decode($item->name, true)['uz'],
                'price' => $item->price,
            ];
        });

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
                'name'  => 'Цена',
                'table' => "C",
            ],
            [
                'name'  => 'Артикул',
                'table' => "D",
            ],
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
                'name'  => 'Цена поставщика',
                'table' => "I",
            ],
        ];

        foreach ($ont_table as $key => $value) {
            $sheet->setCellValue($value['table'].'1', $value['name'])
                ->getStyle($value['table'].'1')
                ->getBorders()
                ->getOutline()
                ->setBorderStyle(Border::BORDER_THICK);
        }
        //        $sheet->getStyle('B2')
        //            ->getBorders()
        //            ->getOutline()
        //            ->setBorderStyle(Border::BORDER_THICK)
        //            ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);


        $validation = $sheet->getCell('B5')->getDataValidation();
        $validation->setType(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_LIST);
        $validation->setErrorStyle(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::STYLE_INFORMATION);
        $validation->setAllowBlank(false);
        $validation->setShowInputMessage(true);
        $validation->setShowErrorMessage(true);
        $validation->setShowDropDown(true);
        $validation->setErrorTitle('Input error');
        $validation->setError('Value is not in list.');
        $validation->setPromptTitle('Pick from list');
        $validation->setPrompt('Please pick a value from the drop-down list.');
        $validation->setFormula1('"Item A,Item B,Item C"');

        //
        //        $product_list = '';
        //
        //        foreach ($products as $product) {
        //            $product_name = $product['name'];
        //            $product_list .= '"'.$product_name.'",';
        //        }
        //
        //        $validation->setFormula1($product_list);


        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;


        //        $spreadsheet = new Spreadsheet();
        //        $sheet       = $sheet->;
        //        $sheet->setCellValue('A1', 'Hello World !');
        //
        //        $writer = new Xlsx($spreadsheet);
        //        $writer->save($filename);
        //
        //        return $filename;
    }


}
