<?php
// vybe-test: php/arrays/array_values_and_keys_reindexing
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

$a = ['x' => 1, 5 => 2, 9 => 3];
$vals = array_values($a);
$keys = array_keys($a);
echo implode(',', $vals) . '|';
echo implode(',', $keys);

__vybe_check(ob_get_clean(), "1,2,3|x,5,9");
