<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_key_first_last_empty_array
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs

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

$a = [];
echo (array_key_first($a) === null ? 'first-null' : 'first-set') . '|';
echo (array_key_last($a) === null ? 'last-null' : 'last-set');

__vybe_check(ob_get_clean(), "first-null|last-null");
