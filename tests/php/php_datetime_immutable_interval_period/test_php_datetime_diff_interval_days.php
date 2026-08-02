<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_datetime_diff_interval_days
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

$d1 = new DateTimeImmutable("2024-01-01");
$d2 = new DateTimeImmutable("2024-01-15");
$interval = $d1->diff($d2);
echo $interval->format("%r%a days");

__vybe_check(ob_get_clean(), "14 days");
