<?php
// vybe-test: php/filter_var_boolean_null/filter_var_boolean_null_on_failure
// origin: languages/php/tests/php/test_filter_var_boolean_null.rs

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

echo filter_var("maybe", FILTER_VALIDATE_BOOLEAN, FILTER_NULL_ON_FAILURE) === null ? "null" : "fail";

__vybe_check(ob_get_clean(), "null");
