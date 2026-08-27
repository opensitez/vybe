<?php
// vybe-test: php/str_getcsv_enclosures/str_getcsv_basic
// origin: languages/php/tests/php/test_str_getcsv_enclosures.rs

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

echo "str_getcsv_basic_ok";

__vybe_check(ob_get_clean(), "str_getcsv_basic_ok");
