<?php
// vybe-test: php/string_builtins_extended/quoted_printable_handles_unicode_roundtrip
// origin: languages/php/tests/php/test_string_builtins_extended.rs

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

echo quoted_printable_encode("Hello");
echo "|";
echo quoted_printable_decode("=48=65=6C=6C=6F");

__vybe_check(ob_get_clean(), "Hello|Hello");
