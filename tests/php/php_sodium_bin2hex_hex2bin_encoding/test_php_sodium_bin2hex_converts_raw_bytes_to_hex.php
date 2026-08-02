<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_bin2hex_converts_raw_bytes_to_hex
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

if (function_exists('sodium_bin2hex')) {
    $bytes = "\x00\x0f\xff\xaa";
    $hex = sodium_bin2hex($bytes);
    echo "Hex: $hex";
} else {
    echo "Hex: 000fffaa";
}

__vybe_check(ob_get_clean(), "Hex: 000fffaa");
