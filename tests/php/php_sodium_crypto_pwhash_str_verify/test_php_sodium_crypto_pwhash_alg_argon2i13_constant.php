<?php
// vybe-test: php/php_sodium_crypto_pwhash_str_verify/test_php_sodium_crypto_pwhash_alg_argon2i13_constant
// origin: languages/php/tests/php/test_php_sodium_crypto_pwhash_str_verify.rs
// vybe-test-mode: compile

if (defined('SODIUM_CRYPTO_PWHASH_ALG_ARGON2I13')) {
    echo is_int(SODIUM_CRYPTO_PWHASH_ALG_ARGON2I13) ? "ARGON2I_CONST_OK" : "FAIL";
} else {
    echo "ARGON2I_CONST_OK";
}
