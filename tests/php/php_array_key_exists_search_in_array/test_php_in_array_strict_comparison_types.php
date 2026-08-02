<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_in_array_strict_comparison_types
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs

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

$arr = [1, "2", 3];
$loose = in_array(2, $arr);
$strict = in_array(2, $arr, true);
echo ($loose ? "L1" : "L0") . " " . ($strict ? "S1" : "S0");

__vybe_check(ob_get_clean(), "L1 S0");
