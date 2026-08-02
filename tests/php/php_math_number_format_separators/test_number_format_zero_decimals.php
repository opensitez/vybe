<?php
// vybe-test: php/php_math_number_format_separators/test_number_format_zero_decimals
// origin: languages/php/tests/php/test_php_math_number_format_separators.rs

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

echo number_format(1234567.89, 0, '.', ','), "\n";

__vybe_check(ob_get_clean(), "1,234,568");
