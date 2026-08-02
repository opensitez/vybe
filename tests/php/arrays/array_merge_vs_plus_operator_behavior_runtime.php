<?php
// vybe-test: php/arrays/array_merge_vs_plus_operator_behavior_runtime
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

$a = ['x' => 1, 2 => 3];
$b = ['x' => 9, 2 => 8, 3 => 7];
$m = array_merge($a, $b);
$u = $a + $b;
echo $m['x'] . '|';
echo $m[2] . '|';
echo $m[3] . '|';
echo $u[2];

__vybe_check(ob_get_clean(), "1|3|7|3");
