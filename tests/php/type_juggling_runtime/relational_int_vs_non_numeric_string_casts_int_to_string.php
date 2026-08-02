<?php
// vybe-test: php/type_juggling_runtime/relational_int_vs_non_numeric_string_casts_int_to_string
// origin: languages/php/tests/php/test_type_juggling_runtime.rs

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

echo (5 < "abc") ? 'lt' : 'ge';

__vybe_check(ob_get_clean(), "lt");
