<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_array_pad_positive_negative_length
// origin: languages/php/tests/php/test_php_array_fill_range_pad_column.rs

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

$input = [12, 10, 9];
$padded_right = array_pad($input, 5, 0);
$padded_left = array_pad($input, -5, -1);
echo implode(",", $padded_right) . " | " . implode(",", $padded_left);

__vybe_check(ob_get_clean(), "12,10,9,0,0 | -1,-1,12,10,9");
