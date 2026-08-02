<?php
// vybe-test: php/php_string_strspn_offset_length/test_strspn_offset_and_max_length
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

echo strspn('foo123456bar', '0123456789', 3, 2), "\n";

__vybe_check(ob_get_clean(), "2");
