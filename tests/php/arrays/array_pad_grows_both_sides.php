<?php
// vybe-test: php/arrays/array_pad_grows_both_sides
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

$a = [1, 2];
$left = array_pad($a, -4, 0);
$right = array_pad($a, 4, 9);
ksort($left);
ksort($right);
echo implode(',', $left) . '|';
echo implode(',', $right);

__vybe_check(ob_get_clean(), "0,0,1,2|1,2,9,9");
