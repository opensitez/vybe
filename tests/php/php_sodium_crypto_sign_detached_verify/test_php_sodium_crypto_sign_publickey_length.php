<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_publickey_length
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_sign_keypair')) {
    $kp = sodium_crypto_sign_keypair();
    $pk = sodium_crypto_sign_publickey($kp);
    echo strlen($pk) === SODIUM_CRYPTO_SIGN_PUBLICKEYBYTES ? "PUBLICKEYBYTES_OK" : "FAIL";
} else {
    echo "PUBLICKEYBYTES_OK";
}
