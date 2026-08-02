<?php
// vybe-test: php/array_spread_string_keys/spread_with_non_contiguous_int_keys_and_assoc_overlap
// origin: languages/php/tests/php/test_array_spread_string_keys.rs

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

$left = [1 => 'a', 3 => 'b'];
$right = ['x' => 100, 2 => 'c'];
$result = [...$left, ...$right];
echo $result[0] . '|' . $result[1] . '|' . $result[2];

__vybe_check(ob_get_clean(), "a|b|c");
