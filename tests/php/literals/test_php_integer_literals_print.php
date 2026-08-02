<?php
// vybe-test: php/literals/test_php_integer_literals_print
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

echo 127;
echo '\n';
echo 0x2A;
echo '\n';
echo 0o77;
echo '\n';
echo 0b1010;
echo '\n';
echo 1_000_000;

__vybe_check(ob_get_clean(), "127\n42\n63\n10\n1000000");
