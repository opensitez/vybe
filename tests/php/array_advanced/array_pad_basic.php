<?php
// vybe-test: php/array_advanced/array_pad_basic
// origin: languages/php/tests/php/test_array_advanced.rs

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

$a = [1, 2, 3];
$padded = array_pad($a, 6, 0);
echo implode(",", $padded);
$left = array_pad($a, -6, 0);
echo implode(",", $left);

__vybe_check(ob_get_clean(), "1,2,3,0,0,00,0,0,1,2,3");
