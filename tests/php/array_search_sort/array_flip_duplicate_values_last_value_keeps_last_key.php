<?php
// vybe-test: php/array_search_sort/array_flip_duplicate_values_last_value_keeps_last_key
// origin: languages/php/tests/php/test_array_search_sort.rs

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

$f = array_flip(['a', 'b', 'a', 'c']);
echo $f['a'];

__vybe_check(ob_get_clean(), "2");
