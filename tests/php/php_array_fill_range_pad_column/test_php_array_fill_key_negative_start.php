<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_array_fill_key_negative_start
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

$a = array_fill(-2, 3, "z");
echo $a[-2] . "|" . $a[-1] . "|" . $a[0];

__vybe_check(ob_get_clean(), "z|z|z");
