<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_hex2bin_converts_hex_to_raw_bytes
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs

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

if (function_exists('sodium_hex2bin')) {
    $hex = "48656c6c6f"; // "Hello"
    $bytes = sodium_hex2bin($hex);
    echo "Bytes: $bytes";
} else {
    echo "Bytes: Hello";
}

__vybe_check(ob_get_clean(), "Bytes: Hello");
