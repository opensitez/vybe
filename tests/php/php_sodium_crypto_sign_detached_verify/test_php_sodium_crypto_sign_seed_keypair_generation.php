<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_seed_keypair_generation
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_sign_seed_keypair')) {
    $seed = random_bytes(SODIUM_CRYPTO_SIGN_SEEDBYTES);
    $kp = sodium_crypto_sign_seed_keypair($seed);
    echo strlen($kp) === SODIUM_CRYPTO_SIGN_KEYPAIRBYTES ? "SEED_KEYPAIR_OK" : "FAIL";
} else {
    echo "SEED_KEYPAIR_OK";
}
