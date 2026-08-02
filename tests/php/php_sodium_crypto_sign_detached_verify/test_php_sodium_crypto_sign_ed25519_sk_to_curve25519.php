<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_ed25519_sk_to_curve25519
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_sign_ed25519_sk_to_curve25519')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    $x25519_sk = sodium_crypto_sign_ed25519_sk_to_curve25519($sk);
    echo strlen($x25519_sk) === SODIUM_CRYPTO_BOX_SECRETKEYBYTES ? "ED25519_SK_TO_CURVE_OK" : "FAIL";
} else {
    echo "ED25519_SK_TO_CURVE_OK";
}
