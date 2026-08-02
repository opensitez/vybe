<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_secretkey_length
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_sign_keypair')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    echo strlen($sk) === SODIUM_CRYPTO_SIGN_SECRETKEYBYTES ? "SECRETKEYBYTES_OK" : "FAIL";
} else {
    echo "SECRETKEYBYTES_OK";
}
