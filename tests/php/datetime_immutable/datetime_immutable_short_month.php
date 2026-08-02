<?php
// vybe-test: php/datetime_immutable/datetime_immutable_short_month
// origin: languages/php/tests/php/test_datetime_immutable.rs

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
$d = new DateTimeImmutable('2024-03-15', new DateTimeZone('UTC'));
echo $d->format('M');

__vybe_check(ob_get_clean(), "Mar");
