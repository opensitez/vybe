<?php
// vybe-test: php/array_functions/array_diff_assoc_missing_and_extra
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

$a = ['a' => 1, 'b' => 2, 'c' => 2];
$b = ['b' => 2, 'd' => 4];
$d = array_diff_assoc($a, $b);
ksort($d);
echo implode(',', array_keys($d)) . '|' . $d['a'];

__vybe_check(ob_get_clean(), "a,c|1");
