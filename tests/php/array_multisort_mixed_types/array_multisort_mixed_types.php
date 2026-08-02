<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_mixed_types
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

$ar1 = ["10", 11, 100, 100, "a"];
$ar2 = [1, 2, "2", 3, 1];
array_multisort($ar1, SORT_ASC, SORT_STRING,
                $ar2, SORT_NUMERIC, SORT_DESC);
echo implode(',', $ar1) . '|' . implode(',', $ar2);

__vybe_check(ob_get_clean(), "10,100,100,11,a|1,3,2,2,1");
