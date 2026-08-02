<?php
// vybe-test: php/arrays/array_unshift_push_pop_order
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

$a = [3, 4];
$len1 = array_unshift($a, 1, 2);
$len2 = array_push($a, 5);
$tail = array_pop($a);
echo $len1 . '|' . $len2 . '|' . $tail . '|';
echo implode(',', $a);

__vybe_check(ob_get_clean(), "4|5|5|1,2,3,4");
