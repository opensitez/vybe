<?php
// vybe-test: php/string_formatting/sprintf_decimal_rounding_and_width_runtime
// origin: languages/php/tests/php/test_string_formatting.rs

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

echo "sprintf_decimal_rounding_and_width_runtime_ok";

__vybe_check(ob_get_clean(), "sprintf_decimal_rounding_and_width_runtime_ok");
