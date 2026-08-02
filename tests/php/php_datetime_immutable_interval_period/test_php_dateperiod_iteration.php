<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_dateperiod_iteration
// origin: languages/php/tests/php/test_php_datetime_immutable_interval_period.rs

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

$start = new DateTimeImmutable("2024-05-01");
$interval = new DateInterval("P1D");
$period = new DatePeriod($start, $interval, 2); // start + 2 occurrences

$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format("Y-m-d");
}
echo implode(", ", $dates);

__vybe_check(ob_get_clean(), "2024-05-01, 2024-05-02, 2024-05-03");
