<?php
// vybe-test: php/php_openssl_cipher_iv_length_lookup/test_php_openssl_random_pseudo_bytes_strong
// origin: languages/php/tests/php/test_php_openssl_cipher_iv_length_lookup.rs
// vybe-test-mode: compile

if (function_exists('openssl_random_pseudo_bytes')) {
    $bytes = openssl_random_pseudo_bytes(16, $strong);
    echo strlen($bytes) === 16 && $strong ? "STRONG_RANDOM_BYTES_OK" : "FAIL";
} else {
    echo "STRONG_RANDOM_BYTES_OK";
}
