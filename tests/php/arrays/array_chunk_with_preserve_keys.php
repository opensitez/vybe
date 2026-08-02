<?php
// vybe-test: php/arrays/array_chunk_with_preserve_keys
// origin: languages/php/tests/php/test_arrays.rs

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

$a = [10 => 'a', 11 => 'b', 12 => 'c', 13 => 'd'];
$chunks = array_chunk($a, 2, true);
$first = array_keys($chunks[0]);
$second = array_keys($chunks[1]);
echo $first[0] . ',' . $first[1] . '|';
echo $second[0] . ',' . $second[1];

__vybe_check(ob_get_clean(), "10,11|12,13");
