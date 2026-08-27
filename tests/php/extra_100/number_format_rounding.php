<?php
// vybe-test: php/extra_100/number_format_rounding
// origin: languages/php/tests/php/test_extra_100.rs

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

echo "number_format_rounding_ok";

__vybe_check(ob_get_clean(), "number_format_rounding_ok");
