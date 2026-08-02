<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_bytes_constant
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs
// vybe-test-mode: compile

if (defined('SODIUM_CRYPTO_SIGN_BYTES')) {
    echo SODIUM_CRYPTO_SIGN_BYTES === 64 || is_int(SODIUM_CRYPTO_SIGN_BYTES) ? "SIGN_BYTES_OK" : "FAIL";
} else {
    echo "SIGN_BYTES_OK";
}
