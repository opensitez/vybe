<?php
// vybe-test: php/array_callbacks/array_walk_recursive_preserves_nested_keys
// origin: languages/php/tests/php/test_array_callbacks.rs

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

$tree = ['a' => ['count' => 1], 'b' => ['count' => 2]];
array_walk_recursive($tree, function(&$value, $key) {
    if (is_numeric($value)) {
        $value += 1;
    }
});
echo $tree['a']['count'];
echo $tree['b']['count'];

__vybe_check(ob_get_clean(), "23");
