<?php
// vybe-test: php/datetime_immutable/datetime_immutable_set_timezone
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
$d = new DateTimeImmutable('2024-01-01 12:00:00', new DateTimeZone('UTC'));
$d2 = $d->setTimezone(new DateTimeZone('Europe/London'));
echo $d2->getTimezone()->getName();

__vybe_check(ob_get_clean(), "Europe/London");
