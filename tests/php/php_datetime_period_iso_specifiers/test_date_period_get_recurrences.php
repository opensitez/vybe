<?php
// vybe-test: php/php_datetime_period_iso_specifiers/test_date_period_get_recurrences
// origin: languages/php/tests/php/test_php_datetime_period_iso_specifiers.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$start = new DateTime('2024-01-01');
$interval = new DateInterval('P1D');
$period = new DatePeriod($start, $interval, 5);
echo $period->getRecurrences(), "\n";

__vybe_check(ob_get_clean(), "5");
