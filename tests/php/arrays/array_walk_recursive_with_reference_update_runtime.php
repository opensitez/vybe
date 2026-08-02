<?php
// vybe-test: php/arrays/array_walk_recursive_with_reference_update_runtime
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

$tree = [
    ['a' => 1],
    ['b' => 2],
];
array_walk_recursive($tree, function (&$value) { $value = $value * 10; });
echo $tree[0]['a'] . '|' . $tree[1]['b'];

__vybe_check(ob_get_clean(), "10|20");
