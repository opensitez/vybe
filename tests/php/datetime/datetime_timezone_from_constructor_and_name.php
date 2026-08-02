<?php
// vybe-test: php/datetime/datetime_timezone_from_constructor_and_name
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

$tz = new DateTimeZone('America/New_York');
$dt = new DateTime('2024-01-02 15:00:00', $tz);
echo $dt->format('e');
echo $dt->getTimezone()->getName();

__vybe_check(ob_get_clean(), "America/New_YorkAmerica/New_York");
