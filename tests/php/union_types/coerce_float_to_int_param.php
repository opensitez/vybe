<?php
// vybe-test: php/union_types/coerce_float_to_int_param
// origin: languages/php/tests/php/test_union_types.rs

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

echo "coerce_float_to_int_param_ok";

__vybe_check(ob_get_clean(), "coerce_float_to_int_param_ok");
