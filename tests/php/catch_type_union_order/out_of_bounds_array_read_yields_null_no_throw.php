<?php
// vybe-test: php/catch_type_union_order/out_of_bounds_array_read_yields_null_no_throw
// origin: languages/php/tests/php/test_catch_type_union_order.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "out_of_bounds_array_read_yields_null_no_throw_ok";

__vybe_check(ob_get_clean(), "out_of_bounds_array_read_yields_null_no_throw_ok");
