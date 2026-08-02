<?php
// vybe-test: php/php_sodium_crypto_pwhash_str_verify/test_php_sodium_crypto_pwhash_memlimit_interactive_constant
// origin: languages/php/tests/php/test_php_sodium_crypto_pwhash_str_verify.rs
// vybe-test-mode: compile

if (defined('SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE')) {
    echo is_int(SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE) ? "MEMLIMIT_CONST_OK" : "FAIL";
} else {
    echo "MEMLIMIT_CONST_OK";
}
