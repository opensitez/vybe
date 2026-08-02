<?php
// vybe-test: php/math_advanced/base_convert_basic
// origin: languages/php/tests/php/test_math_advanced.rs

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

echo base_convert("ff", 16, 10);
echo base_convert("255", 10, 16);
echo base_convert("111", 2, 10);
echo base_convert("10", 10, 2);

__vybe_check(ob_get_clean(), "255ff71010");
