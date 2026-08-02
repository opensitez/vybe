<?php
// vybe-test: php/dateinterval_format_specifiers/dateinterval_format_inverted
// origin: languages/php/tests/php/test_dateinterval_format_specifiers.rs

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

$start = new DateTime('2024-01-01');
$end = new DateTime('2023-01-01');
$diff = $end->diff($start);
echo $diff->invert . '|' . $diff->format('%R%y years');

__vybe_check(ob_get_clean(), "0|+1 years");
