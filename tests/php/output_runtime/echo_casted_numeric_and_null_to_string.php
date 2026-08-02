<?php
// vybe-test: php/output_runtime/echo_casted_numeric_and_null_to_string
// origin: languages/php/tests/php/test_output_runtime.rs

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

echo (string) 12;
echo (string) 3.5;
echo (string) null;
echo (string) false;

__vybe_check(ob_get_clean(), "123.5");
