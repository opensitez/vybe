<?php
// vybe-test: php/literals/test_php_boolean_cast_literals_and_is_numeric
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

echo (int) true;
echo '|';
echo (int) false;
echo '|';
echo is_numeric('1_000') ? 'n' : 'v';
echo '|';
echo is_numeric('1000') ? 'n' : 'v';

__vybe_check(ob_get_clean(), "1|0|v|n");
