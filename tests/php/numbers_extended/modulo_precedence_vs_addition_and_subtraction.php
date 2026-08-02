<?php
// vybe-test: php/numbers_extended/modulo_precedence_vs_addition_and_subtraction
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

echo 10 + 9 % 3;
echo '|';
echo 10 - 9 % 3;
echo '|';
echo (10 + 9) % 3;

__vybe_check(ob_get_clean(), "10|10|1");
