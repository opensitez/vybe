<?php
// vybe-test: php/php_array_chunk_slice_splice_combine/test_php_array_combine_keys_and_values
// origin: languages/php/tests/php/test_php_array_chunk_slice_splice_combine.rs

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

$keys = ["x", "y", "z"];
$values = [1, 2, 3];
$combined = array_combine($keys, $values);
echo $combined["y"];

__vybe_check(ob_get_clean(), "2");
