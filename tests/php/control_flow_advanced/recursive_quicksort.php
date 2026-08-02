<?php
// vybe-test: php/control_flow_advanced/recursive_quicksort
// origin: languages/php/tests/php/test_control_flow_advanced.rs

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

function qsort(array $a): array {
    if (count($a) <= 1) return $a;
    $pivot = $a[0];
    $left  = array_filter(array_slice($a,1), fn($x) => $x <= $pivot);
    $right = array_filter(array_slice($a,1), fn($x) => $x > $pivot);
    return [...qsort(array_values($left)), $pivot, ...qsort(array_values($right))];
}
echo implode(',', qsort([3,6,8,10,1,2,1]));

__vybe_check(ob_get_clean(), "1,1,2,3,6,8,10");
