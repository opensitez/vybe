<?php
// vybe-test: php/array_walk_recursive_by_ref/array_walk_recursive_nested_scalar_collection
// origin: languages/php/tests/php/test_array_walk_recursive_by_ref.rs

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

$a = ['a' => [1, 2, 3], 'b' => ['c' => 4], 'd' => true];
$leaves = [];
array_walk_recursive($a, function($v, $k) use (&$leaves) {
    if (is_scalar($v)) {
        $leaves[] = $k . '=' . (string)$v;
    }
});
echo implode('|', $leaves);

__vybe_check(ob_get_clean(), "0=1|1=2|2=3|c=4|d=1");
