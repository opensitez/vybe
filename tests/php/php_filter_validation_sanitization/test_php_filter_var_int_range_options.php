<?php
// vybe-test: php/php_filter_validation_sanitization/test_php_filter_var_int_range_options
// origin: languages/php/tests/php/test_php_filter_validation_sanitization.rs

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

$val = 42;
$options = ["options" => ["min_range" => 1, "max_range" => 50]];
echo filter_var($val, FILTER_VALIDATE_INT, $options) !== false ? "IN_RANGE" : "OUT_OF_RANGE";

__vybe_check(ob_get_clean(), "IN_RANGE");
