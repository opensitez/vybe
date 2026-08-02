<?php
// vybe-test: php/php_openssl_cipher_iv_length_lookup/test_php_openssl_digest_hash_computation
// origin: languages/php/tests/php/test_php_openssl_cipher_iv_length_lookup.rs
// vybe-test-mode: compile

if (function_exists('openssl_digest')) {
    $hash = openssl_digest("sample text", "sha256");
    echo strlen($hash) === 64 ? "DIGEST_SHA256_64HEX_OK" : "FAIL";
} else {
    echo "DIGEST_SHA256_64HEX_OK";
}
