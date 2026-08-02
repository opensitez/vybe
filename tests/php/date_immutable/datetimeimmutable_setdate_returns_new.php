<?php
// vybe-test: php/date_immutable/datetimeimmutable_setdate_returns_new
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

$d = new DateTimeImmutable('2024-01-01');
$d2 = $d->setDate(2025, 6, 15);
echo $d->format('Y') . ',' . $d2->format('Y-m-d');

__vybe_check(ob_get_clean(), "2024,2025-06-15");
