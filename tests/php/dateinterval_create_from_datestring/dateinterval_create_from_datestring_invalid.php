<?php
// vybe-test: php/dateinterval_create_from_datestring/dateinterval_create_from_datestring_invalid
// origin: languages/php/tests/php/test_dateinterval_create_from_datestring.rs

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

echo "dateinterval_create_from_datestring_invalid_ok";

__vybe_check(ob_get_clean(), "dateinterval_create_from_datestring_invalid_ok");
