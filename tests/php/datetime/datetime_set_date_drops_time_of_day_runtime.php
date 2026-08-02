<?php
// vybe-test: php/datetime/datetime_set_date_drops_time_of_day_runtime
// origin: languages/php/tests/php/test_datetime.rs

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

$dt = new DateTime('2024-01-01 12:34:56');
$dt->setDate(2024, 2, 15);
echo $dt->format('Y-m-d H:i:s');

__vybe_check(ob_get_clean(), "2024-02-15 12:34:56");
