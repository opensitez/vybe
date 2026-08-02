<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_fixed_array_bounds_and_size
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs

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

$arr = new SplFixedArray(3);
$arr[0] = 10;
$arr[1] = 20;
$arr[2] = 30;
echo count($arr) . " | " . $arr[1];

__vybe_check(ob_get_clean(), "3 | 20");
