<?php
// vybe-test: php/datetime/datetime_format_various
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

$dt = new DateTime("2024-12-25 09:30:45");
echo $dt->format("l");
echo $dt->format("F j, Y");
echo $dt->format("g:i A");

__vybe_check(ob_get_clean(), "WednesdayDecember 25, 20249:30 AM");
