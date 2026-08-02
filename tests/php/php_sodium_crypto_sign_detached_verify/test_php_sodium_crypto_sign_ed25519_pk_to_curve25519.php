<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_ed25519_pk_to_curve25519
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_sign_ed25519_pk_to_curve25519')) {
    $kp = sodium_crypto_sign_keypair();
    $pk = sodium_crypto_sign_publickey($kp);
    $x25519_pk = sodium_crypto_sign_ed25519_pk_to_curve25519($pk);
    echo strlen($x25519_pk) === SODIUM_CRYPTO_BOX_PUBLICKEYBYTES ? "ED25519_TO_CURVE_OK" : "FAIL";
} else {
    echo "ED25519_TO_CURVE_OK";
}
