<?php
// vybe-test: php/array_advanced3/array_first_and_last
// origin: languages/php/tests/php/test_array_advanced3.rs

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

$a = [10,20,30,40,50];
echo array_key_first($a) . ':' . $a[array_key_first($a)];
echo ',';
echo array_key_last($a) . ':' . $a[array_key_last($a)];

__vybe_check(ob_get_clean(), "0:10,4:50");
