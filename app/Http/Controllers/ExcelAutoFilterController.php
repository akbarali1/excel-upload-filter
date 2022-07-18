<?php

namespace App\Http\Controllers;

use DateTime;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule;
use PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\NamedRange;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;

/**
 * Created by PhpStorm.
 * Filename: ExcelAutoFilterController.php
 * Project Name: excelfiler.loc
 * Author: Акбарали
 * Date: 18/07/2022
 * Time: 17:08
 * Github: https://github.com/akbarali1
 * Telegram: @akbar_aka
 * E-mail: me@akbarali.uz
 */
class ExcelAutoFilterController extends Controller
{

    public function index()
    {
        return view('excel.index');
    }

    public function download()
    {
        $spreadsheet = new Spreadsheet();

        $file     = $this->generateExcelDropdownDownload();
        $content  = file_get_contents($file);
        $response = response($content, 200)->header('Content-Type', 'application/vnd.ms-excel');
        $response->header('Content-Disposition', 'attachment; filename="'.$file.'"');
        unlink($file);

        return $response;

    }

    public function transpose($value): array
    {
        return [$value];
    }

    public function generateExcelDropdownDownload($file = 'helloWorld.xlsx')
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getProperties()
            ->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPOffice')
            ->setTitle('PhpSpreadsheet Test Document')
            ->setSubject('PhpSpreadsheet Test Document')
            ->setDescription('Test document for PhpSpreadsheet, generated using PHP classes.')
            ->setKeywords('Office PhpSpreadsheet php')
            ->setCategory('Test result file');

        // Add some data
        $continentColumn = 'D';
        $column          = 'F';


        // Set data for dropdowns
        $continents = glob(__DIR__.'/data/continents/*');
        foreach ($continents as $key => $filename) {
            $continent    = pathinfo($filename, PATHINFO_FILENAME);
            $continent    = str_replace(' ', '_', $continent);
            $countries    = file($filename, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
            $countryCount = count($countries);

            // Transpose $countries from a row to a column array

            $countries = array_map('self::transpose', $countries);
            $spreadsheet->getActiveSheet()->fromArray($countries, null, $column.'1');
            $spreadsheet->addNamedRange(
                new NamedRange(
                    $continent,
                    $spreadsheet->getActiveSheet(),
                    '$'.$column.'$1:$'.$column.'$'.$countryCount
                )
            );
            $spreadsheet->getActiveSheet()
                ->getColumnDimension($column)
                ->setVisible(false);

            $spreadsheet->getActiveSheet()
                ->setCellValue($continentColumn.($key + 1), $continent);

            ++$column;
        }

        // Hide the dropdown data
        $spreadsheet->getActiveSheet()
            ->getColumnDimension($continentColumn)
            ->setVisible(false);

        $spreadsheet->addNamedRange(
            new NamedRange(
                'Continents',
                $spreadsheet->getActiveSheet(),
                '$'.$continentColumn.'$1:$'.$continentColumn.'$'.count($continents)
            )
        );

        // Set selection cells
        $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'Continent:');
        $spreadsheet->getActiveSheet()
            ->setCellValue('B1', 'Select continent');
        $spreadsheet->getActiveSheet()
            ->setCellValue('B3', '='.$column. 1);
        $spreadsheet->getActiveSheet()
            ->setCellValue('B3', 'Select country');
        $spreadsheet->getActiveSheet()
            ->getStyle('A1:A3')
            ->getFont()->setBold(true);

        // Set linked validators
        $validation = $spreadsheet->getActiveSheet()
            ->getCell('B1')
            ->getDataValidation();
        $validation->setType(DataValidation::TYPE_LIST)
            ->setErrorStyle(DataValidation::STYLE_INFORMATION)
            ->setAllowBlank(false)
            ->setShowInputMessage(true)
            ->setShowErrorMessage(true)
            ->setShowDropDown(true)
            ->setErrorTitle('Input error')
            ->setError('Continent is not in the list.')
            ->setPromptTitle('Pick from the list')
            ->setPrompt('Please pick a continent from the drop-down list.')
            ->setFormula1('=Continents');

        $spreadsheet->getActiveSheet()
            ->setCellValue('A3', 'Country:');
        $spreadsheet->getActiveSheet()
            ->getStyle('A3')
            ->getFont()->setBold(true);

        $validation = $spreadsheet->getActiveSheet()
            ->getCell('B3')
            ->getDataValidation();
        $validation->setType(DataValidation::TYPE_LIST)
            ->setErrorStyle(DataValidation::STYLE_INFORMATION)
            ->setAllowBlank(false)
            ->setShowInputMessage(true)
            ->setShowErrorMessage(true)
            ->setShowDropDown(true)
            ->setErrorTitle('Input error')
            ->setError('Country is not in the list.')
            ->setPromptTitle('Pick from the list')
            ->setPrompt('Please pick a country from the drop-down list.')
            ->setFormula1('=INDIRECT($B$1)');

        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(12);
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(30);


        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;
    }

    public function generateExcelAutoFilterSelectionDisplayDownload($file = 'helloWorld.xlsx')
    {
        $spreadsheet = new Spreadsheet();

        $spreadsheet->getProperties()->setCreator('Maarten Balliauw')
            ->setLastModifiedBy('Maarten Balliauw')
            ->setTitle('PhpSpreadsheet Test Document')
            ->setSubject('PhpSpreadsheet Test Document')
            ->setDescription('Test document for PhpSpreadsheet, generated using PHP classes.')
            ->setKeywords('office PhpSpreadsheet php')
            ->setCategory('Test result file');

        // Create the worksheet
        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'Financial Year')
            ->setCellValue('B1', 'Financial Period')
            ->setCellValue('C1', 'Country')
            ->setCellValue('D1', 'Date')
            ->setCellValue('E1', 'Sales Value')
            ->setCellValue('F1', 'Expenditure');
        $dateTime  = new DateTime();
        $startYear = $endYear = $currentYear = (int)$dateTime->format('Y');
        --$startYear;
        ++$endYear;

        $years     = range($startYear, $endYear);
        $periods   = range(1, 12);
        $countries = [
            'United States',
            'UK',
            'France',
            'Germany',
            'Italy',
            'Spain',
            'Portugal',
            'Japan',
        ];

        $row = 2;
        foreach ($years as $year) {
            foreach ($periods as $period) {
                foreach ($countries as $country) {
                    $dateString = sprintf('%04d-%02d-01T00:00:00', $year, $period);
                    $dateTime   = new DateTime($dateString);
                    $endDays    = (int)$dateTime->format('t');
                    for ($i = 1; $i <= $endDays; ++$i) {
                        $eDate               = Date::formattedPHPToExcel(
                            $year,
                            $period,
                            $i
                        );
                        $value               = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        $salesValue          = $invoiceValue = null;
                        $incomeOrExpenditure = mt_rand(-1, 1);
                        if ($incomeOrExpenditure == -1) {
                            $expenditure = mt_rand(-1000, -500) * (1 + (mt_rand(-1, 1) / 4));
                            $income      = null;
                        } elseif ($incomeOrExpenditure == 1) {
                            $expenditure = mt_rand(-1000, -500) * (1 + (mt_rand(-1, 1) / 4));
                            $income      = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        } else {
                            $expenditure = null;
                            $income      = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        }
                        $dataArray = [
                            $year,
                            $period,
                            $country,
                            $eDate,
                            $income,
                            $expenditure,
                        ];
                        $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A'.$row++);
                    }
                }
            }
        }
        --$row;

        // Set styling
        $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(12.5);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(10.5);
        $spreadsheet->getActiveSheet()->getStyle('D2:D'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_YYYYMMDD2);
        $spreadsheet->getActiveSheet()->getStyle('E2:F'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(14);
        $spreadsheet->getActiveSheet()->freezePane('A2');

        // Set autofilter range
        // Always include the complete filter range!
        // Excel does support setting only the caption
        // row, but that's not a best practise...
        $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        // Set active filters
        $autoFilter = $spreadsheet->getActiveSheet()->getAutoFilter();
        // Filter the Country column on a filter value of countries beginning with the letter U (or Japan)
        //     We use * as a wildcard, so specify as U* and using a wildcard requires customFilter
        $autoFilter->getColumn('C')
            ->setFilterType(Column::AUTOFILTER_FILTERTYPE_CUSTOMFILTER)
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                'u*'
            )
            ->setRuleType(Rule::AUTOFILTER_RULETYPE_CUSTOMFILTER);
        $autoFilter->getColumn('C')
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                'japan'
            )
            ->setRuleType(Rule::AUTOFILTER_RULETYPE_CUSTOMFILTER);
        // Filter the Date column on a filter value of the last day of every period of the current year
        // We us a dateGroup ruletype for this, although it is still a standard filter
        foreach ($periods as $period) {
            $dateString = sprintf('%04d-%02d-01T00:00:00', $currentYear, $period);
            $dateTime   = new DateTime($dateString);
            $endDate    = (int)$dateTime->format('t');

            $autoFilter->getColumn('D')
                ->setFilterType(Column::AUTOFILTER_FILTERTYPE_FILTER)
                ->createRule()
                ->setRule(
                    Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                    [
                        'year'  => $currentYear,
                        'month' => $period,
                        'day'   => $endDate,
                    ]
                )
                ->setRuleType(Rule::AUTOFILTER_RULETYPE_DATEGROUP);
        }
        // Display only sales values that are blank
        //     Standard filter, operator equals, and value of NULL
        $autoFilter->getColumn('E')
            ->setFilterType(Column::AUTOFILTER_FILTERTYPE_FILTER)
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                ''
            );

        // Execute filtering
        $autoFilter->showHideRows();

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);

        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;
    }

    public function generateExcelAutoFilterSelectionTwoDownload($file = 'helloWorld.xlsx')
    {
        $spreadsheet = new Spreadsheet();

        $spreadsheet->getProperties()->setCreator('Maarten Balliauw')
            ->setLastModifiedBy('Maarten Balliauw')
            ->setTitle('PhpSpreadsheet Test Document')
            ->setSubject('PhpSpreadsheet Test Document')
            ->setDescription('Test document for PhpSpreadsheet, generated using PHP classes.')
            ->setKeywords('office PhpSpreadsheet php')
            ->setCategory('Test result file');

        // Create the worksheet
        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'Financial Year')
            ->setCellValue('B1', 'Financial Period')
            ->setCellValue('C1', 'Country')
            ->setCellValue('D1', 'Date')
            ->setCellValue('E1', 'Sales Value')
            ->setCellValue('F1', 'Expenditure');
        $dateTime  = new DateTime();
        $startYear = $endYear = $currentYear = (int)$dateTime->format('Y');
        --$startYear;
        ++$endYear;

        $years     = range($startYear, $endYear);
        $periods   = range(1, 12);
        $countries = [
            'United States',
            'UK',
            'France',
            'Germany',
            'Italy',
            'Spain',
            'Portugal',
            'Japan',
        ];

        $row = 2;
        foreach ($years as $year) {
            foreach ($periods as $period) {
                foreach ($countries as $country) {
                    $dateString = sprintf('%04d-%02d-01T00:00:00', $year, $period);
                    $dateTime   = new DateTime($dateString);
                    $endDays    = (int)$dateTime->format('t');
                    for ($i = 1; $i <= $endDays; ++$i) {
                        $eDate               = Date::formattedPHPToExcel(
                            $year,
                            $period,
                            $i
                        );
                        $value               = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        $salesValue          = $invoiceValue = null;
                        $incomeOrExpenditure = mt_rand(-1, 1);
                        if ($incomeOrExpenditure == -1) {
                            $expenditure = mt_rand(-1000, -500) * (1 + (mt_rand(-1, 1) / 4));
                            $income      = null;
                        } elseif ($incomeOrExpenditure == 1) {
                            $expenditure = mt_rand(-1000, -500) * (1 + (mt_rand(-1, 1) / 4));
                            $income      = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        } else {
                            $expenditure = null;
                            $income      = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        }
                        $dataArray = [
                            $year,
                            $period,
                            $country,
                            $eDate,
                            $income,
                            $expenditure,
                        ];
                        $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A'.$row++);
                    }
                }
            }
        }
        --$row;

        // Set styling
        $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(12.5);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(10.5);
        $spreadsheet->getActiveSheet()->getStyle('D2:D'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_YYYYMMDD2);
        $spreadsheet->getActiveSheet()->getStyle('E2:F'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(14);
        $spreadsheet->getActiveSheet()->freezePane('A2');

        // Set autofilter range
        // Always include the complete filter range!
        // Excel does support setting only the caption
        // row, but that's not a best practise...
        $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        // Set active filters
        $autoFilter = $spreadsheet->getActiveSheet()->getAutoFilter();
        // Filter the Country column on a filter value of Germany
        // As it's just a simple value filter, we can use FILTERTYPE_FILTER
        $autoFilter->getColumn('C')
            ->setFilterType(Column::AUTOFILTER_FILTERTYPE_FILTER)
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                'Germany'
            );
        // Filter the Date column on a filter value of the year to date
        $autoFilter->getColumn('D')
            ->setFilterType(Column::AUTOFILTER_FILTERTYPE_DYNAMICFILTER)
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                null,
                Rule::AUTOFILTER_RULETYPE_DYNAMIC_YEARTODATE
            )
            ->setRuleType(Rule::AUTOFILTER_RULETYPE_DYNAMICFILTER);
        // Display only sales values that are between 400 and 600
        $autoFilter->getColumn('E')
            ->setFilterType(Column::AUTOFILTER_FILTERTYPE_CUSTOMFILTER)
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_GREATERTHANOREQUAL,
                400
            )
            ->setRuleType(Rule::AUTOFILTER_RULETYPE_CUSTOMFILTER);
        $autoFilter->getColumn('E')
            ->setJoin(Column::AUTOFILTER_COLUMN_JOIN_AND)
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_LESSTHANOREQUAL,
                600
            )
            ->setRuleType(Rule::AUTOFILTER_RULETYPE_CUSTOMFILTER);


        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;
    }

    public function generateExcelAutoFilterSelectionDownload($file = 'helloWorld.xlsx')
    {
        $spreadsheet = new Spreadsheet();


        $spreadsheet->getProperties()->setCreator('Maarten Balliauw')
            ->setLastModifiedBy('Maarten Balliauw')
            ->setTitle('PhpSpreadsheet Test Document')
            ->setSubject('PhpSpreadsheet Test Document')
            ->setDescription('Test document for PhpSpreadsheet, generated using PHP classes.')
            ->setKeywords('office PhpSpreadsheet php')
            ->setCategory('Test result file');


        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'Financial Year')
            ->setCellValue('B1', 'Financial Period')
            ->setCellValue('C1', 'Country')
            ->setCellValue('D1', 'Date')
            ->setCellValue('E1', 'Sales Value')
            ->setCellValue('F1', 'Expenditure');
        $dateTime  = new DateTime();
        $startYear = $endYear = $currentYear = (int)$dateTime->format('Y');
        --$startYear;
        ++$endYear;

        $years     = range($startYear, $endYear);
        $periods   = range(1, 12);
        $countries = [
            'United States',
            'UK',
            'France',
            'Germany',
            'Italy',
            'Spain',
            'Portugal',
            'Japan',
        ];

        $row = 2;
        foreach ($years as $year) {
            foreach ($periods as $period) {
                foreach ($countries as $country) {
                    $dateString = sprintf('%04d-%02d-01T00:00:00', $year, $period);
                    $dateTime   = new DateTime($dateString);
                    $endDays    = (int)$dateTime->format('t');
                    for ($i = 1; $i <= $endDays; ++$i) {
                        $eDate               = Date::formattedPHPToExcel(
                            $year,
                            $period,
                            $i
                        );
                        $value               = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        $salesValue          = $invoiceValue = null;
                        $incomeOrExpenditure = mt_rand(-1, 1);
                        if ($incomeOrExpenditure == -1) {
                            $expenditure = mt_rand(-1000, -500) * (1 + (mt_rand(-1, 1) / 4));
                            $income      = null;
                        } elseif ($incomeOrExpenditure == 1) {
                            $expenditure = mt_rand(-1000, -500) * (1 + (mt_rand(-1, 1) / 4));
                            $income      = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        } else {
                            $expenditure = null;
                            $income      = mt_rand(500, 1000) * (1 + (mt_rand(-1, 1) / 4));
                        }
                        $dataArray = [
                            $year,
                            $period,
                            $country,
                            $eDate,
                            $income,
                            $expenditure,
                        ];
                        $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A'.$row++);
                    }
                }
            }
        }
        --$row;

        // Set styling
        $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(12.5);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(10.5);
        $spreadsheet->getActiveSheet()->getStyle('D2:D'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_YYYYMMDD2);
        $spreadsheet->getActiveSheet()->getStyle('E2:F'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(14);
        $spreadsheet->getActiveSheet()->freezePane('A2');

        // Set autofilter range
        // Always include the complete filter range!
        // Excel does support setting only the caption
        // row, but that's not a best practise...
        $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        // Set active filters
        $autoFilter = $spreadsheet->getActiveSheet()->getAutoFilter();
        // Filter the Country column on a filter value of countries beginning with the letter U (or Japan)
        //     We use * as a wildcard, so specify as U* and using a wildcard requires customFilter
        $autoFilter->getColumn('C')
            ->setFilterType(Column::AUTOFILTER_FILTERTYPE_CUSTOMFILTER)
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                'u*'
            )
            ->setRuleType(Rule::AUTOFILTER_RULETYPE_CUSTOMFILTER);
        $autoFilter->getColumn('C')
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                'japan'
            )
            ->setRuleType(Rule::AUTOFILTER_RULETYPE_CUSTOMFILTER);
        // Filter the Date column on a filter value of the last day of every period of the current year
        // We us a dateGroup ruletype for this, although it is still a standard filter
        foreach ($periods as $period) {
            $dateString = sprintf('%04d-%02d-01T00:00:00', $currentYear, $period);
            $dateTime   = new DateTime($dateString);
            $endDate    = (int)$dateTime->format('t');

            $autoFilter->getColumn('D')
                ->setFilterType(Column::AUTOFILTER_FILTERTYPE_FILTER)
                ->createRule()
                ->setRule(
                    Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                    [
                        'year'  => $currentYear,
                        'month' => $period,
                        'day'   => $endDate,
                    ]
                )
                ->setRuleType(Rule::AUTOFILTER_RULETYPE_DATEGROUP);
        }
        // Display only sales values that are blank
        //     Standard filter, operator equals, and value of NULL
        $autoFilter->getColumn('E')
            ->setFilterType(Column::AUTOFILTER_FILTERTYPE_FILTER)
            ->createRule()
            ->setRule(
                Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                ''
            );


        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;
    }

    public function generateExcelDynamicDateDownload($file = 'helloWorld.xlsx')
    {
        function createSheet(Spreadsheet $spreadsheet, string $rule): void
        {
            $sheet = $spreadsheet->createSheet();
            $sheet->setTitle($rule);
            $sheet->getCell('A1')->setValue('Date');
            $row        = 1;
            $date       = new DateTime();
            $year       = (int)$date->format('Y');
            $month      = (int)$date->format('m');
            $day        = (int)$date->format('d');
            $yearMinus2 = $year - 2;
            $sheet->getCell('B1')->setValue("=DATE($year, $month, $day)");
            // Each day for two weeks before today through 2 weeks after
            for ($dayOffset = -14; $dayOffset < 14; ++$dayOffset) {
                ++$row;
                $sheet->getCell("A$row")->setValue("=B1+($dayOffset)");
            }
            // First and last day of each month, starting with January 2 years before,
            // through December 2 years after.
            for ($monthOffset = 0; $monthOffset < 48; ++$monthOffset) {
                ++$row;
                $sheet->getCell("A$row")->setValue("=DATE($yearMinus2, $monthOffset, 1)");
                ++$row;
                $sheet->getCell("A$row")->setValue("=DATE($yearMinus2, $monthOffset + 1, 0)");
            }
            $sheet->getStyle("A2:A$row")->getNumberFormat()->setFormatCode('yyyy-mm-dd');
            $sheet->getStyle('B1')->getNumberFormat()->setFormatCode('yyyy-mm-dd');
            $sheet->getColumnDimension('A')->setAutoSize(true);
            $sheet->getColumnDimension('B')->setAutoSize(true);
            $autoFilter = $spreadsheet->getActiveSheet()->getAutoFilter();
            $autoFilter->setRange("A1:A$row");
            $columnFilter = $autoFilter->getColumn('A');
            $columnFilter->setFilterType(Column::AUTOFILTER_FILTERTYPE_DYNAMICFILTER);
            $columnFilter->createRule()
                ->setRule(
                    Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                    '',
                    $rule
                )
                ->setRuleType(Rule::AUTOFILTER_RULETYPE_DYNAMICFILTER);
            $sheet->setSelectedCell('B1');
        }

        $spreadsheet = new Spreadsheet();
        $spreadsheet->getProperties()->setCreator('Owen Leibman')
            ->setLastModifiedBy('Owen Leibman')
            ->setTitle('PhpSpreadsheet Test Document')
            ->setSubject('PhpSpreadsheet Test Document')
            ->setDescription('Test document for PhpSpreadsheet, generated using PHP classes.')
            ->setKeywords('office PhpSpreadsheet php')
            ->setCategory('Test result file');

        $ruleNames = [
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTMONTH,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTQUARTER,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTWEEK,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTYEAR,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTMONTH,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTQUARTER,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTWEEK,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTYEAR,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_THISMONTH,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_THISQUARTER,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_THISWEEK,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_THISYEAR,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_TODAY,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_TOMORROW,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_YEARTODATE,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_YESTERDAY,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_2,
            Rule::AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_3,
        ];

        // Create the worksheets
        foreach ($ruleNames as $ruleName) {
            createSheet($spreadsheet, $ruleName);
        }
        $spreadsheet->removeSheetByIndex(0);
        $spreadsheet->setActiveSheetIndex(0);

        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;
    }

    public function generateExcelAutoFilterDownload($file = 'helloWorld.xlsx')
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getProperties()->setCreator('Maarten Balliauw')
            ->setLastModifiedBy('Maarten Balliauw')
            ->setTitle('PhpSpreadsheet Test Document')
            ->setSubject('PhpSpreadsheet Test Document')
            ->setDescription('Test document for PhpSpreadsheet, generated using PHP classes.')
            ->setKeywords('office PhpSpreadsheet php')
            ->setCategory('Test result file');

        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'Year')
            ->setCellValue('B1', 'Quarter')
            ->setCellValue('C1', 'Country')
            ->setCellValue('D1', 'Sales');

        $dataArray = [
            ['2010', 'Q1', 'United States', 790],
            ['2010', 'Q2', 'United States', 730],
            ['2010', 'Q3', 'United States', 860],
            ['2010', 'Q4', 'United States', 850],
            ['2011', 'Q1', 'United States', 800],
            ['2011', 'Q2', 'United States', 700],
            ['2011', 'Q3', 'United States', 900],
            ['2011', 'Q4', 'United States', 950],
            ['2010', 'Q1', 'Belgium', 380],
            ['2010', 'Q2', 'Belgium', 390],
            ['2010', 'Q3', 'Belgium', 420],
            ['2010', 'Q4', 'Belgium', 460],
            ['2011', 'Q1', 'Belgium', 400],
            ['2011', 'Q2', 'Belgium', 350],
            ['2011', 'Q3', 'Belgium', 450],
            ['2011', 'Q4', 'Belgium', 500],
            ['2010', 'Q1', 'UK', 690],
            ['2010', 'Q2', 'UK', 610],
            ['2010', 'Q3', 'UK', 620],
            ['2010', 'Q4', 'UK', 600],
            ['2011', 'Q1', 'UK', 720],
            ['2011', 'Q2', 'UK', 650],
            ['2011', 'Q3', 'UK', 580],
            ['2011', 'Q4', 'UK', 510],
            ['2010', 'Q1', 'France', 510],
            ['2010', 'Q2', 'France', 490],
            ['2010', 'Q3', 'France', 460],
            ['2010', 'Q4', 'France', 590],
            ['2011', 'Q1', 'France', 620],
            ['2011', 'Q2', 'France', 650],
            ['2011', 'Q3', 'France', 415],
            ['2011', 'Q4', 'France', 570],
            ['2010', 'Q1', 'Germany', 720],
            ['2010', 'Q2', 'Germany', 680],
            ['2010', 'Q3', 'Germany', 640],
            ['2010', 'Q4', 'Germany', 660],
            ['2011', 'Q1', 'Germany', 680],
            ['2011', 'Q2', 'Germany', 620],
            ['2011', 'Q3', 'Germany', 710],
            ['2011', 'Q4', 'Germany', 690],
            ['2010', 'Q1', 'Spain', 510],
            ['2010', 'Q2', 'Spain', 490],
            ['2010', 'Q3', 'Spain', 470],
            ['2010', 'Q4', 'Spain', 420],
            ['2011', 'Q1', 'Spain', 460],
            ['2011', 'Q2', 'Spain', 390],
            ['2011', 'Q3', 'Spain', 430],
            ['2011', 'Q4', 'Spain', 415],
            ['2010', 'Q1', 'Italy', 440],
            ['2010', 'Q2', 'Italy', 410],
            ['2010', 'Q3', 'Italy', 420],
            ['2010', 'Q4', 'Italy', 450],
            ['2011', 'Q1', 'Italy', 430],
            ['2011', 'Q2', 'Italy', 370],
            ['2011', 'Q3', 'Italy', 350],
            ['2011', 'Q4', 'Italy', 335],
        ];
        $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A2');

        $spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return $file;
    }

}
