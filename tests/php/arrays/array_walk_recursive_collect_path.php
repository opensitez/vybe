<?php
// vybe-test: php/arrays/array_walk_recursive_collect_path
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

$tree = ['a' => ['x' => 1], 'b' => ['y' => 2]];
$out = [];
array_walk_recursive($tree, function($v, $k) use (&$out) { $out[] = "$k=$v"; });
echo implode('|', $out);

__vybe_check(ob_get_clean(), "x=1|y=2");
