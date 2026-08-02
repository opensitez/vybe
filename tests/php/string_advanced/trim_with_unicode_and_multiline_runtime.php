<?php
// vybe-test: php/string_advanced/trim_with_unicode_and_multiline_runtime
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

echo trim(" \n\tHello World\t\n");
echo "\n";
echo trim("xxHelloxx", "x");
echo "\n";
echo trim("xyxxyx", "xy");

__vybe_check(ob_get_clean(), "Hello World\nHello\n");
