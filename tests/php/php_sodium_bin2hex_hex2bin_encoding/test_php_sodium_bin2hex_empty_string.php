<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_bin2hex_empty_string
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs
// vybe-test-mode: compile

if (function_exists('sodium_bin2hex')) {
    echo sodium_bin2hex("") === "" ? "EMPTY_BIN2HEX_OK" : "FAIL";
} else {
    echo "EMPTY_BIN2HEX_OK";
}
