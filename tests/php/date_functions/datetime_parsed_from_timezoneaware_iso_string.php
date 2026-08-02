<?php
// vybe-test: php/date_functions/datetime_parsed_from_timezoneaware_iso_string
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
$dt = new DateTime('2024-07-01T10:15:00+02:00');
echo $dt->format('Y-m-d H:i');
echo '|';
echo $dt->getTimezone()->getName();

__vybe_check(ob_get_clean(), "2024-07-01 10:15|+02:00");
