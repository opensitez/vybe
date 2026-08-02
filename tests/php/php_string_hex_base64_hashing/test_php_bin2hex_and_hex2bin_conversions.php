<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_bin2hex_and_hex2bin_conversions
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs

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

$binary = "\x48\x65\x6c\x6c\x6f"; // "Hello"
$hex = bin2hex($binary);
$restored = hex2bin($hex);

echo "$hex | $restored";

__vybe_check(ob_get_clean(), "48656c6c6f | Hello");
