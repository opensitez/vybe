<?php
// vybe-test: php/date_functions/date_create_from_format_with_timezone_object
// origin: languages/php/tests/php/test_date_functions.rs

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
$tz = new DateTimeZone('America/Los_Angeles');
$dt = date_create_from_format('Y-m-d H:i', '2024-11-05 01:30', $tz);
echo $dt->format('e');
echo '|';
echo $dt->format('H:i');

__vybe_check(ob_get_clean(), "America/Los_Angeles|01:30");
