<?php
// vybe-test: php/php_string_strspn_offset_length/test_strspn_zero_length_subject_returns_zero
// origin: languages/php/tests/php/test_php_string_strspn_offset_length.rs

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

echo strspn('', 'a', 0), "\n";
echo strspn('', 'a', 0, 10), "\n";

__vybe_check(ob_get_clean(), "0\n0");
