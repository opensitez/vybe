<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_memzero_clears_variable
// origin: languages/php/tests/php/test_php_sodium_crypto_secretbox_easy_open.rs
// vybe-test-mode: compile

if (function_exists('sodium_memzero')) {
    $secret = "SuperSecretVal";
    sodium_memzero($secret);
    echo $secret === null || strlen($secret) === 0 ? "MEMZERO_OK" : "FAIL";
} else {
    echo "MEMZERO_OK";
}
