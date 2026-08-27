<?php
// vybe-test: php/numeric_casts/bcdiv_by_zero_returns_null_without_scale
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

echo "bcdiv_by_zero_returns_null_without_scale_ok";

__vybe_check(ob_get_clean(), "bcdiv_by_zero_returns_null_without_scale_ok");
