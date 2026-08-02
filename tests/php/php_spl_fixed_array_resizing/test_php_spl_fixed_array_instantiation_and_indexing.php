<?php
// vybe-test: php/php_spl_fixed_array_resizing/test_php_spl_fixed_array_instantiation_and_indexing
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

$arr = new SplFixedArray(3);
$arr[0] = "zero";
$arr[1] = "one";
$arr[2] = "two";

echo "Size=" . $arr->getSize() . " Val1=" . $arr[1];

__vybe_check(ob_get_clean(), "Size=3 Val1=one");
