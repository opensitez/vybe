<?php
// vybe-test: php/php_datetime_period_iso_specifiers/test_date_period_iso_get_start_end_dates
// origin: languages/php/tests/php/test_php_datetime_period_iso_specifiers.rs

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

echo "test_date_period_iso_get_start_end_dates_ok";

__vybe_check(ob_get_clean(), "test_date_period_iso_get_start_end_dates_ok");
