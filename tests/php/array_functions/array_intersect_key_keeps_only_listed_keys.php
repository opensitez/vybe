<?php
// vybe-test: php/array_functions/array_intersect_key_keeps_only_listed_keys
// origin: languages/php/tests/php/test_array_functions.rs

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

$data = ['a' => 1, 'b' => 2, 'c' => 3];
$only = array_intersect_key($data, array_flip(['a', 'c']));
echo implode(',', array_keys($only));

__vybe_check(ob_get_clean(), "a,c");
