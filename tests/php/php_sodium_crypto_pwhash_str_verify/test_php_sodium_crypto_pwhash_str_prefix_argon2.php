<?php
// vybe-test: php/php_sodium_crypto_pwhash_str_verify/test_php_sodium_crypto_pwhash_str_prefix_argon2
// origin: languages/php/tests/php/test_php_sodium_crypto_pwhash_str_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_pwhash_str')) {
    $hash = sodium_crypto_pwhash_str(
        "pass",
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    echo str_starts_with($hash, "$argon2") ? "ARGON2_PREFIX_HASH_OK" : "FAIL";
} else {
    echo "ARGON2_PREFIX_HASH_OK";
}
