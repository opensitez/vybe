<?php
// vybe-test: php/arrays/array_intersection_by_key_values
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

$a = ['a' => 1, 'b' => 2, 'c' => 3];
$b = ['a' => 99, 'c' => 33];
$both = array_intersect_key($a, $b);
$all = array_intersect_assoc($a, array_merge($b, ['c' => 3]));
ksort($both);
ksort($all);
echo implode(',', array_keys($both)) . '|';
echo implode(',', $all);

__vybe_check(ob_get_clean(), "a,c|c");
