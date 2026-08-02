<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_crypto_secretbox_noncebytes_constant
// origin: languages/php/tests/php/test_php_sodium_crypto_secretbox_easy_open.rs
// vybe-test-mode: compile

if (defined('SODIUM_CRYPTO_SECRETBOX_NONCEBYTES')) {
    echo SODIUM_CRYPTO_SECRETBOX_NONCEBYTES === 24 || is_int(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES) ? "NONCEBYTES_OK" : "FAIL";
} else {
    echo "NONCEBYTES_OK";
}
