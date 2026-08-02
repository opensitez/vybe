<?php
// vybe-test: php/php_string_strspn_strcspn_masks/test_strcspn_negative_offset
// origin: languages/php/tests/php/test_php_string_strspn_strcspn_masks.rs

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

echo strcspn('a1b2c3', '123', -4), "\n";

__vybe_check(ob_get_clean(), "1");
