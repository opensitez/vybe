<?php
// vybe-test: php/array_functions/array_rand_two_keys_sorted
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

$r = array_rand(['a' => 1, 'b' => 2, 'c' => 3, 'd' => 4], 3);
sort($r);
echo implode(',', $r);

__vybe_check(ob_get_clean(), "a,b,c");
