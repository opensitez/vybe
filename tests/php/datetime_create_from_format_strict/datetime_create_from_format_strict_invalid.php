<?php
// vybe-test: php/datetime_create_from_format_strict/datetime_create_from_format_strict_invalid
// origin: languages/php/tests/php/test_datetime_create_from_format_strict.rs

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

echo "datetime_create_from_format_strict_invalid_ok";

__vybe_check(ob_get_clean(), "datetime_create_from_format_strict_invalid_ok");
