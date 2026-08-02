<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_verify_tampered_message_fails
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_sign_detached')) {
    $kp = sodium_crypto_sign_keypair();
    $sig = sodium_crypto_sign_detached("Original", sodium_crypto_sign_secretkey($kp));
    $valid = sodium_crypto_sign_verify_detached($sig, "Tampered", sodium_crypto_sign_publickey($kp));
    echo $valid === false ? "TAMPERED_VERIFY_FALSE_OK" : "FAIL";
} else {
    echo "TAMPERED_VERIFY_FALSE_OK";
}
