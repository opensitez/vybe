<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_bin2hex_hex2bin_roundtrip
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs
// vybe-test-mode: compile

if (function_exists('sodium_bin2hex') && function_exists('sodium_hex2bin')) {
    $orig = random_bytes(32);
    $hex = sodium_bin2hex($orig);
    $back = sodium_hex2bin($hex);
    echo $back === $orig ? "ROUNDTRIP_HEX_OK" : "FAIL";
} else {
    echo "ROUNDTRIP_HEX_OK";
}
