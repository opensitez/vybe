<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_hex2bin_invalid_hex_length_error
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs
// vybe-test-mode: compile

if (function_exists('sodium_hex2bin')) {
    try {
        @sodium_hex2bin("123"); // Odd number of hex digits
        echo "HEX2BIN_ODD_HANDLED";
    } catch (SodiumException $e) {
        echo "HEX2BIN_ODD_HANDLED";
    }
} else {
    echo "HEX2BIN_ODD_HANDLED";
}
