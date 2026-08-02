<?php
// vybe-test: php/datetime/mktime_basic
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

$ts1 = mktime(0, 0, 0, 1, 1, 2020);
$ts2 = mktime(0, 0, 0, 1, 2, 2020);
echo $ts2 - $ts1;

__vybe_check(ob_get_clean(), "86400");
