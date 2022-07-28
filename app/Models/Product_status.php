<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class Product_status extends Model
{

    protected $table = "product_status";

    protected $fillable = [
        "id",
        "sku",
        'serial_number',
        "status",
    ];


}
