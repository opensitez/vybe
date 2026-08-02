<?php
// vybe-test: php/datetime/date_timezone_abbreviation_lookup_runtime
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
$abbr = $tz->getName();
$time = new DateTime('2024-07-01 12:00:00', $tz);
echo $abbr . '|' . $time->format('e') . '|' . $time->getOffset();

__vybe_check(ob_get_clean(), "America/New_York|America/New_York|-14400");
