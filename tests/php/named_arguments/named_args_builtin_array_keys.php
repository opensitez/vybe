<?php
// vybe-test: php/named_arguments/named_args_builtin_array_keys
// origin: languages/php/tests/php/test_named_arguments.rs

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

$map = ['a' => 1, 'b' => 2, 'c' => 1, 'd' => 3];
$keys = array_keys(array: $map, filter_value: 1);
echo implode(',', $keys) . "\n";

__vybe_check(ob_get_clean(), "a,c");
