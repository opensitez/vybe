<?php
// vybe-test: php/filter_validation/filter_sanitize_number_int_strips_letters
// origin: languages/php/tests/php/test_filter_validation.rs

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

echo filter_var('a1b2c3', FILTER_SANITIZE_NUMBER_INT);

__vybe_check(ob_get_clean(), "123");
