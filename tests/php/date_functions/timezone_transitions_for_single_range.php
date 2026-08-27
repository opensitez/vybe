<?php
// vybe-test: php/date_functions/timezone_transitions_for_single_range
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

echo "timezone_transitions_for_single_range_ok";

__vybe_check(ob_get_clean(), "timezone_transitions_for_single_range_ok");
