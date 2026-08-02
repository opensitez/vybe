<?php
// vybe-test: php/arrays/array_fill_keys_ordered_numeric_and_string_keys
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

$map = array_fill_keys([0, 1, 'x', 'y'], 7);
ksort($map);
echo isset($map[0]) ? '0' : 'x';
echo isset($map['x']) ? '|x' : '|no';
echo count($map);

__vybe_check(ob_get_clean(), "0|x4");
