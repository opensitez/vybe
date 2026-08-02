<?php
// vybe-test: php/literals/test_php_float_literals_print
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

echo 3.5;
echo '\n';
echo 1e2;
echo '\n';
echo 1.2e-3;
echo '\n';
echo (1 + 2.5);

__vybe_check(ob_get_clean(), "3.5\n100\n0.0012\n3.5");
