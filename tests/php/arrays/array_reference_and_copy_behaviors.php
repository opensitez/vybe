<?php
// vybe-test: php/arrays/array_reference_and_copy_behaviors
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

$source = [1, 2, 3];
$copy = $source;
$source[0] = 9;
echo $copy[0] . "\n";
$alias = &$source;
$alias[1] = 11;
echo $source[1];

__vybe_check(ob_get_clean(), "1\n11");
