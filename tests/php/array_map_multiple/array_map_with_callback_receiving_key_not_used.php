<?php
// vybe-test: php/array_map_multiple/array_map_with_callback_receiving_key_not_used
// origin: languages/php/tests/php/test_array_map_multiple.rs

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

$input = ['a' => 1, 'b' => 2, 'c' => 3];
$doubled = array_map(fn($v) => $v * 2, $input);
echo $doubled['a'] . ',' . $doubled['b'] . ',' . $doubled['c'];

__vybe_check(ob_get_clean(), "2,4,6");
