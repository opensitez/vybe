<?php
// vybe-test: php/arrays/array_combine_and_map_runtime
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

$k = ['a', 'b', 'c'];
$v = [1, 2, 3];
$c = array_combine($k, $v);
$m = array_map(fn($x) => $x * $x, $c);
echo $m['a'], "\n";
echo $m['b'], "\n";
echo $m['c'];

__vybe_check(ob_get_clean(), "1\n4\n9");
