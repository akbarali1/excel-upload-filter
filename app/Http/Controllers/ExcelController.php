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

        $this->excelHeaderGenerate($sheet);
        $this->excelBodyGenerate($sheet);

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

    public function excelBodyGenerate($sheet)
    {
        $products = DB::table('products')->whereNull('deleted_at')
            ->orderBy('sku')
            ->select('id', 'name')
            ->get()->transform(function ($item) {
                $name = json_decode($item->name, true);

                return [
                    'id'   => $item->id,
                    'name' => $name['uz'],
                ];
            })->toArray();

        $i = 1;
        $g = 0;
        foreach ($products as $key => $value) {
            $i++;
            $g++;

            $sheet->getStyle('A'.$i.':J'.$i)->applyFromArray([
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                    ],
                ],
            ]);

            $sheet->setCellValue('A'.$i, $g.')')
                ->setCellValue('B'.$i, $value['name'])
                ->setCellValue('C'.$i, $value['id'])
                ->setCellValue('D'.$i, '0')
                ->setCellValue('E'.$i, '0')
                ->setCellValue('F'.$i, '0')
                ->setCellValue('G'.$i, '0')
                ->setCellValue('H'.$i, '0')
                ->setCellValue('I'.$i, '0')
                ->setCellValue('J'.$i, '0');

        }

        //        $validation = $sheet->getCell('B5')->getDataValidation();
        //        $validation->setType(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_LIST);
        //        $validation->setErrorStyle(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::STYLE_INFORMATION);
        //        $validation->setAllowBlank(false);
        //        $validation->setShowInputMessage(true);
        //        $validation->setShowErrorMessage(true);
        //        $validation->setShowDropDown(true);
        //        $validation->setErrorTitle('Input error');
        //        $validation->setError('Value is not in list.');
        //        $validation->setPromptTitle('Pick from list');
        //        $validation->setPrompt('Please pick a value from the drop-down list.');
        //        $validation->setFormula1('"Item A,Item B,Item C"');

        return $sheet;

    }


}
