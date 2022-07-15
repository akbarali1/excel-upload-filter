<?php

namespace App\Http\Controllers;

use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

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
        $i           = 3;

        $providersList = DB::table('products')->limit(100)->get()->transform(function ($item) {
            return [
                'id'    => $item->id,
                'name'  => json_decode($item->name, true)['uz'],
                'price' => $item->price,
            ];
        });

        foreach ($providersList as $provider) {
            $sheet->setCellValue('A'.$i, $provider['id']);
            $sheet->setCellValue('B'.$i, $provider['name']);
            $i++;
        }
        $sheet->getProtection()->setSheet(true);
        $nbOfProvider = $sheet->getHighestRow('A');
        for ($j = 3; $j < $nbOfProvider; $j++) {
            $dropdownlist = $sheet->getCell('B'.$j)->getDataValidation();
            $dropdownlist->setType(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_LIST)
                ->setAllowBlank(true)
                ->setShowDropDown(true)
                ->setPrompt('Choose the provider')
                ->setFormula1('=\'PROVIDERS\'!$B$3:$B$'.$nbOfProvider);
            $idProvider = $sheet->getCell('C'.$j);
            /*$idProvider->setValue('=VLOOKUP(\'PROVIDERS\'!$B$3:$B$'.$nbOfProvider.',B');*/
        }
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
