<?php
// vybe-test: php/array_functions/array_rand_singleton_returns_scalar
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

$k = array_rand(['a' => 1, 'b' => 2, 'c' => 3], 1);
echo is_string($k) ? 'scalar' : 'array';

__vybe_check(ob_get_clean(), "scalar");
