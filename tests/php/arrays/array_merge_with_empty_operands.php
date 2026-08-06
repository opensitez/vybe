<?php
// vybe-test: php/arrays/array_merge_with_empty_operands
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

$a = [1, 2];
$b = array_merge($a, []);
$c = array_merge([], $a);
echo count($b) . '|';
echo count($c) . '|';
echo $b[0] . $c[1];

__vybe_check(ob_get_clean(), "2|2|12");
