<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_hex2bin_ignore_chars_parameter
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs
// vybe-test-mode: compile

if (function_exists('sodium_hex2bin')) {
    $hexWithSpaces = "48 65 6c 6c 6f";
    $bytes = sodium_hex2bin($hexWithSpaces, " ");
    echo $bytes === "Hello" ? "HEX2BIN_IGNORE_SPACES_OK" : "FAIL";
} else {
    echo "HEX2BIN_IGNORE_SPACES_OK";
}
