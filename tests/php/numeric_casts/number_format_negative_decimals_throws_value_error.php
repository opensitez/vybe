<?php
// vybe-test: php/numeric_casts/number_format_negative_decimals_throws_value_error
// origin: languages/php/tests/php/test_numeric_casts.rs

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

echo "number_format_negative_decimals_throws_value_error_ok";

__vybe_check(ob_get_clean(), "number_format_negative_decimals_throws_value_error_ok");
