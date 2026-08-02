<?php
// vybe-test: php/array_creation/array_merge_recursive_combines_nested
// origin: languages/php/tests/php/test_array_creation.rs

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

$a = ['x' => ['a' => 1]];
$b = ['x' => ['b' => 2]];
echo json_encode(array_merge_recursive($a, $b));

__vybe_check(ob_get_clean(), "{\"x\":{\"a\":1,\"b\":2}}");
