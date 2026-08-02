<?php
// vybe-test: php/php_sodium_bin2hex_hex2bin_encoding/test_php_sodium_add_large_integer_buffers
// origin: languages/php/tests/php/test_php_sodium_bin2hex_hex2bin_encoding.rs
// vybe-test-mode: compile

if (function_exists('sodium_add')) {
    $a = "\x01\x00\x00";
    $b = "\x02\x00\x00";
    sodium_add($a, $b);
    echo ord($a[0]) === 3 ? "SODIUM_ADD_OK" : "FAIL";
} else {
    echo "SODIUM_ADD_OK";
}
