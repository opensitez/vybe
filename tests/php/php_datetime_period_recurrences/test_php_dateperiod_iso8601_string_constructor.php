<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_iso8601_string_constructor
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs

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

// Repeat 3 times every 2 days starting from 2024-05-01
$period = new DatePeriod("R3/2024-05-01T00:00:00Z/P2D");
$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format("Y-m-d");
}
echo implode(", ", $dates);

__vybe_check(ob_get_clean(), "2024-05-01, 2024-05-03, 2024-05-05, 2024-05-07");
