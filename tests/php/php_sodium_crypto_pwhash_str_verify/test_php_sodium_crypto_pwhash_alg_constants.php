<?php
// vybe-test: php/php_sodium_crypto_pwhash_str_verify/test_php_sodium_crypto_pwhash_alg_constants
// origin: languages/php/tests/php/test_php_sodium_crypto_pwhash_str_verify.rs
// vybe-test-mode: compile

if (defined('SODIUM_CRYPTO_PWHASH_ALG_ARGON2ID13')) {
    echo is_int(SODIUM_CRYPTO_PWHASH_ALG_ARGON2ID13) ? "ARGON2ID_CONST_OK" : "FAIL";
} else {
    echo "ARGON2ID_CONST_OK";
}
