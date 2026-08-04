<?php
// vybe-test: php/datetime/strtotime_relative
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

$base = strtotime("2024-06-15");
$next = strtotime("+7 days", $base);
echo date("Y-m-d", $next);
$prev = strtotime("-1 month", $base);
echo date("Y-m-d", $prev);

__vybe_check(ob_get_clean(), "2024-06-222024-05-15");
