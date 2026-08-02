<?php
// vybe-test: php/datetime/datetime_set_timezone_affects_date_output_runtime
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

date_default_timezone_set('UTC');
$utc = new DateTime('2024-01-01 12:00:00');
date_default_timezone_set('America/Los_Angeles');
$la = new DateTime('2024-01-01 12:00:00');
echo $utc->format('Y-m-d H:i');
echo '|';
echo $la->format('Y-m-d H:i');

__vybe_check(ob_get_clean(), "2024-01-01 12:00|2024-01-01 12:00");
