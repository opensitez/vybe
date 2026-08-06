<?php
// vybe-test: php/array_functions/array_diff_recursive_with_nested_array
// origin: languages/php/tests/php/test_array_functions.rs

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

$a = ['a' => ['x' => 1], 'b' => ['y' => 2]];
$b = ['a' => ['x' => 1]];
// array_diff() takes only arrays — a flag as argument #3 is a TypeError, and
// array_diff compares elements as STRINGS, so nested arrays all collapse to
// "Array". Comparing nested values needs array_udiff with a real comparator.
$d = array_udiff($a, $b, fn($p, $q) => strcmp(json_encode($p), json_encode($q)));
echo json_encode($d);

__vybe_check(ob_get_clean(), "{\"b\":{\"y\":2}}");
