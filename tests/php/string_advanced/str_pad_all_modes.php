<?php
// vybe-test: php/string_advanced/str_pad_all_modes
// origin: languages/php/tests/php/test_string_advanced.rs

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

echo str_pad("42", 5, "0", STR_PAD_LEFT);
echo "\n";
echo str_pad("hi", 10, "-");
echo "\n";
echo str_pad("x", 5, "AB", STR_PAD_BOTH);
echo "\n";

__vybe_check(ob_get_clean(), "00042\nhi--------\nABxAB");
