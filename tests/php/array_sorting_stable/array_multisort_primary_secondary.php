<?php
// vybe-test: php/array_sorting_stable/array_multisort_primary_secondary
// origin: languages/php/tests/php/test_array_sorting_stable.rs

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

$a = [3,1,3,1,2];
$b = ['e','d','c','b','a'];
array_multisort($a, SORT_ASC, $b, SORT_ASC);
echo implode(',', $a) . '|' . implode(',', $b);

__vybe_check(ob_get_clean(), "1,1,2,3,3|b,d,a,c,e");
