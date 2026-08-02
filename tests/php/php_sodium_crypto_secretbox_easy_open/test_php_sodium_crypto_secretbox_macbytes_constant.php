<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_crypto_secretbox_macbytes_constant
// origin: languages/php/tests/php/test_php_sodium_crypto_secretbox_easy_open.rs
// vybe-test-mode: compile

if (defined('SODIUM_CRYPTO_SECRETBOX_MACBYTES')) {
    echo SODIUM_CRYPTO_SECRETBOX_MACBYTES === 16 || is_int(SODIUM_CRYPTO_SECRETBOX_MACBYTES) ? "MACBYTES_OK" : "FAIL";
} else {
    echo "MACBYTES_OK";
}
