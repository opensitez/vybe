<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_out_of_bounds_exception
// origin: languages/php/tests/php/test_php_spl_fixed_array_resizing.rs

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

$arr = new SplFixedArray(2);
try {
    $val = $arr[5];
} catch (RuntimeException $e) {
    echo "OUT_OF_BOUNDS_EX";
}

__vybe_check(ob_get_clean(), "OUT_OF_BOUNDS_EX");
