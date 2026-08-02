<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_regular_and_nat_combined
// origin: languages/php/tests/php/test_array_multisort_mixed_types.rs

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

$a = ["10", "2", "1", "3"];
$b = ["y", "x", "z", "w"];
array_multisort($a, SORT_ASC, SORT_STRING, $b, SORT_DESC, SORT_NUMERIC);
echo implode(',', $a) . "|" . implode(',', $b);

__vybe_check(ob_get_clean(), "1,10,2,3|z,y,x,w");
