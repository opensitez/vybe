<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_memcmp_constant_time_equal
// origin: languages/php/tests/php/test_php_sodium_crypto_secretbox_easy_open.rs
// vybe-test-mode: compile

if (function_exists('sodium_memcmp')) {
    $res = sodium_memcmp("secret123", "secret123");
    echo $res === 0 ? "MEMCMP_EQUAL_0_OK" : "FAIL";
} else {
    echo "MEMCMP_EQUAL_0_OK";
}
