<?php
// vybe-test: php/array_column_advanced/array_key_exists_with_null_value_and_missing_key
// origin: languages/php/tests/php/test_array_column_advanced.rs

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

$row = ['a' => null];
echo array_key_exists('a', $row) ? 'a_yes' : 'a_no';
echo '|';
echo array_key_exists('b', $row) ? 'b_yes' : 'b_no';

__vybe_check(ob_get_clean(), "a_yes|b_no");
