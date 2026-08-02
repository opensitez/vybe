<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_reverse_preserving_keys
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

$input = ["a" => 1, "b" => 2, "c" => 3];
$reversed = array_reverse($input, preserve_keys: true);
echo implode(",", array_keys($reversed));

__vybe_check(ob_get_clean(), "c,b,a");
