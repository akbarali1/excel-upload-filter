<?php

use App\Http\Controllers\ExcelController;
use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::view('/', 'welcome')->name('welcome');

Route::controller(ExcelController::class)->prefix('excel')->name('excel.')->group(function () {
    Route::get('/', 'index')->name('index');
    Route::get("/download", "download")->name("download");
});

Route::controller(\App\Http\Controllers\ExcelProductController::class)->prefix('excel_prod')->name('excel_prod.')->group(function () {
    Route::get('/', 'index')->name('index');
    Route::get("/download", "download")->name("download");
});
