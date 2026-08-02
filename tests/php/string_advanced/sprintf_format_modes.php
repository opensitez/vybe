<?php
// vybe-test: php/string_advanced/sprintf_format_modes
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

echo sprintf("%x", 255);
echo "\n";
echo sprintf("%X", 255);
echo "\n";
echo sprintf("%o", 8);
echo "\n";
echo sprintf("%b", 10);
echo "\n";
echo sprintf("%e", 123456.789);
echo "\n";

__vybe_check(ob_get_clean(), "ff\nFF\n10\n1010\n1.234568e+5");
