<?php
// vybe-test: php/literals/test_php_numeric_string_and_float_literals
// origin: languages/php/tests/php/test_literals.rs

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

echo 0b111;
echo '\n';
echo 012;
echo '\n';
echo 0x1a;
echo '\n';
echo 1e-3;
echo '\n';
echo 3.25e2;

__vybe_check(ob_get_clean(), "7\n10\n26\n0.001\n325");
