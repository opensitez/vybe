<?php
// vybe-test: php/string_encoding/urldecode_respects_plus_as_space_runtime
// origin: languages/php/tests/php/test_string_encoding.rs

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

echo urldecode("a+b+c");
echo "\n";
echo rawurldecode("a%2Bb%2Bc");

__vybe_check(ob_get_clean(), "a b c\na+b+c");
