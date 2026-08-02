<?php
// vybe-test: php/array_functions/array_fill_mixed_value_types
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

$a = array_fill(1, 4, ['x' => 1]);
$a[1]['x'] = 2;
echo $a[1]['x'] . '|' . $a[2]['x'];

__vybe_check(ob_get_clean(), "2|1");
