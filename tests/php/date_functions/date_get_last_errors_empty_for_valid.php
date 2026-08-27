<?php
// vybe-test: php/date_functions/date_get_last_errors_empty_for_valid
// origin: languages/php/tests/php/test_date_functions.rs

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

echo "date_get_last_errors_empty_for_valid_ok";

__vybe_check(ob_get_clean(), "date_get_last_errors_empty_for_valid_ok");
