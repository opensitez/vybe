<?php
// vybe-test: php/date_immutable/datetimeimmutable_settime_returns_new
// origin: languages/php/tests/php/test_date_immutable.rs

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

$d = new DateTimeImmutable('2024-01-01 00:00:00');
$d2 = $d->setTime(12, 30, 45);
echo $d2->format('H:i:s');

__vybe_check(ob_get_clean(), "12:30:45");
