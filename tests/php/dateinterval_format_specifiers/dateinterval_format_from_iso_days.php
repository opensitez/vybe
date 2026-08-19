<?php
// vybe-test: php/dateinterval_format_specifiers/dateinterval_format_from_iso_days
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

$i = new DateInterval('P15D');
echo $i->format('%R%a');

__vybe_check(ob_get_clean(), "+(unknown)");
