<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_keypair_from_secretkey_and_publickey
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_sign_keypair_from_secretkey_and_publickey')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    $pk = sodium_crypto_sign_publickey($kp);
    $reconstructed = sodium_crypto_sign_keypair_from_secretkey_and_publickey($sk, $pk);
    echo strlen($reconstructed) === SODIUM_CRYPTO_SIGN_KEYPAIRBYTES ? "RECONSTRUCTED_KEYPAIR_OK" : "FAIL";
} else {
    echo "RECONSTRUCTED_KEYPAIR_OK";
}
