<?php
// vybe-test: php/string_advanced/soundex_basic
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

echo soundex("Robert");
echo "\n";
echo soundex("Rupert");
echo "\n";
echo soundex("Robert") == soundex("Rupert") ? "match" : "no match";
echo "\n";

__vybe_check(ob_get_clean(), "R163\nR163\nmatch");
