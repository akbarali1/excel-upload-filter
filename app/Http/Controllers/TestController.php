<?php

namespace App\Http\Controllers;


use App\Models\Product_status;

/**
 * Created by PhpStorm.
 * Filename: TestController.php
 * Project Name: excelfiler.loc
 * Author: Акбарали
 * Date: 28/07/2022
 * Time: 7:06 PM
 * Github: https://github.com/akbarali1
 * Telegram: @akbar_aka
 * E-mail: me@akbarali.uz
 */
class TestController extends Controller
{

    public function index()
    {
        $zone = '[{"Bertoua Zone 3":[{"code":"BFT50C","product_quantity":"4","zone_name":"Bertoua Zone 3"},{"code":"MNY65C","product_quantity":"10","zone_name":"Bertoua Zone 3"},{"code":"JAP65C","product_quantity":"7","zone_name":"Bertoua Zone 3"},{"code":"MNY65C","product_quantity":"4","zone_name":"Bertoua Zone 3"},{"code":"JAP65C","product_quantity":"3","zone_name":"Bertoua Zone 3"},{"code":"BFT50C","product_quantity":"3","zone_name":"Bertoua Zone 3"},{"code":"JAP65C","product_quantity":"2","zone_name":"Bertoua Zone 3"}],"Bertoua Zone 1":[{"code":"MNY65C","product_quantity":"5","zone_name":"Bertoua Zone 1"},{"code":"JAP65C","product_quantity":"3","zone_name":"Bertoua Zone 1"},{"code":"BFT50C","product_quantity":"7","zone_name":"Bertoua Zone 1"},{"code":"MNY65C","product_quantity":"15","zone_name":"Bertoua Zone 1"},{"code":"JAP65C","product_quantity":"5","zone_name":"Bertoua Zone 1"},{"code":"BFT50C","product_quantity":"4","zone_name":"Bertoua Zone 1"}]}]';
        $zone = json_decode($zone, true);

        $zone2 = '[{"Bertoua Zone 3":[{"code":"BFT50C","product_quantity":"7","zone_name":"Bertoua Zone 3"},{"code":"MNY65C","product_quantity":"14","zone_name":"Bertoua Zone 3"},{"code":"JAP65C","product_quantity":"12","zone_name":"Bertoua Zone 3"}],"Bertoua Zone 1":[{"code":"MNY65C","product_quantity":"20","zone_name":"Bertoua Zone 1"},{"code":"JAP65C","product_quantity":"8","zone_name":"Bertoua Zone 1"},{"code":"BFT50C","product_quantity":"11","zone_name":"Bertoua Zone 1"}]}]';
        $zone2 = json_decode($zone2, true);

        $result = [];
        foreach ($zone[0] as $key => $element) {
            foreach ($element as $value) {
                if (isset($result[$key][$value['code']])) {
                    $result[0][$key][$value['code']]['product_quantity'] += $value['product_quantity'];
                } else {
                    $result[0][$key][$value['code']] = $value;
                }
            }
        }

        echo '<pre>';
        echo print_r($result);
        echo '</pre>';
        echo '<hr>';
        echo '<pre>';
        echo print_r($zone2);
        echo '</pre>';

    }

    public function dasilva()
    {

        $list_of_products_status = Product_status::query()->distinct()->where('sku', 'XOX44JUR8AA')->latest('serial_number')->groupBy('serial_number')->get()->pluck('serial_number');
        dd($list_of_products_status);

        $list_unique_serials = $list_of_products_status->pluck('serial_number')->toArray();

        $products = collect([]);
        foreach ($list_unique_serials as $serial) {
            // get the latest status of each product
            $product    = Product_status::query()->where('serial_number', $serial)
                ->latest()->first();
            $products[] = $product;
        }
    }


}
