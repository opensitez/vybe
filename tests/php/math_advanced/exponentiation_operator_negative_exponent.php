<?php
// vybe-test: php/math_advanced/exponentiation_operator_negative_exponent
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

echo 2 ** 3;
echo "\n";
echo 2 ** 0;
echo "\n";
echo 2 ** -1;
echo "\n";
echo 9 ** 0.5;

__vybe_check(ob_get_clean(), "8\n1\n0.5\n3");
