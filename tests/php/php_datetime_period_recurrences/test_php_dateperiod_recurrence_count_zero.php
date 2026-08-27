<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_recurrence_count_zero
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs

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

echo "test_php_dateperiod_recurrence_count_zero_ok";

__vybe_check(ob_get_clean(), "test_php_dateperiod_recurrence_count_zero_ok");
