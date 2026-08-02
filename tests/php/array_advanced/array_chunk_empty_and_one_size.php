<?php
// vybe-test: php/array_advanced/array_chunk_empty_and_one_size
// origin: languages/php/tests/php/test_array_advanced.rs

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

$empty = array_chunk([], 3);
echo count($empty);
$single = array_chunk([1], 3, true);
echo count($single);
echo array_key_first($single[0]);
echo $single[0][0];

__vybe_check(ob_get_clean(), "0101");
