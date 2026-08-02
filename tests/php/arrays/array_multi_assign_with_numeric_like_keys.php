<?php
// vybe-test: php/arrays/array_multi_assign_with_numeric_like_keys
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

$a = [];
$a[] = 'first';
$a['1'] = 'string1';
$a[1.9] = 'float1';
echo count($a) . '|';
echo $a[1] . '|';
echo $a[1.9];

__vybe_check(ob_get_clean(), "2|string1|string1");
