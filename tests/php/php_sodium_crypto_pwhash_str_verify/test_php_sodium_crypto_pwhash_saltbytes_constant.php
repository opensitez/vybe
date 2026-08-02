<?php
// vybe-test: php/php_sodium_crypto_pwhash_str_verify/test_php_sodium_crypto_pwhash_saltbytes_constant
// origin: languages/php/tests/php/test_php_sodium_crypto_pwhash_str_verify.rs
// vybe-test-mode: compile

if (defined('SODIUM_CRYPTO_PWHASH_SALTBYTES')) {
    echo SODIUM_CRYPTO_PWHASH_SALTBYTES === 16 || is_int(SODIUM_CRYPTO_PWHASH_SALTBYTES) ? "SALTBYTES_16_OK" : "FAIL";
} else {
    echo "SALTBYTES_16_OK";
}
