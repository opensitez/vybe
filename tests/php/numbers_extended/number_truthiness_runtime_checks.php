<?php
// vybe-test: php/numbers_extended/number_truthiness_runtime_checks
// origin: languages/php/tests/php/test_numbers_extended.rs

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

echo (0 == false) ? 'T' : 'F';
echo (0 === false) ? 'T' : 'F';
echo (0.0 == false) ? 'T' : 'F';
echo ('' == false) ? 'T' : 'F';
echo ('0' == false) ? 'T' : 'F';

__vybe_check(ob_get_clean(), "TFTTT");
